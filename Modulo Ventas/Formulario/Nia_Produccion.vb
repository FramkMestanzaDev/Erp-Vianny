Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class Nia_Produccion
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
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
    Dim Rsr2, Rsr20, Rsr21, Rsr22, Rsr23, Rsr233, Rsr24, Rsr2333 As SqlDataReader

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

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
    Private Sub Nia_Produccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox4.Text = DateTimePicker1.Value.Month
        TextBox4.ReadOnly = False
        Select Case TextBox4.Text.Length
            Case "1" : TextBox4.Text = "0" & DateTimePicker1.Value.Month
            Case "2" : TextBox4.Text = TextBox4.Text

        End Select
        TextBox4.Select()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

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

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim p, co1 As Integer
        p = DataGridView1.Rows.Count
        co1 = 0
        For i1 = 0 To p - 1
            If Trim(DataGridView1.Rows(i1).Cells(9).Value) = "" Then
                co1 = co1 + 1
            End If
        Next

        If co1 > 0 Then
            MsgBox("ES OBLIGATORIO EL INGRESO DE LA ORDEN DE SERVICIO, FALTA AGREGAR LA ORDEN A UNA FILA")
        Else


            If Me.ValidateChildren() And Trim(TextBox7.Text) = String.Empty Or Trim(TextBox8.Text) = String.Empty Or Trim(TextBox4.Text) = String.Empty Or Trim(TextBox5.Text) = String.Empty Then
                MsgBox("FALTAN INGREAR ALGUNO CAMPOS OBLIGATORIOS")
            Else
                abrir()
                Dim ccom, doc As String

                If ComboBox4.Text = "GUIA" Then
                    doc = "009"
                Else
                    If ComboBox4.Text = "OTRO" Then
                        doc = "000"
                    Else
                        If ComboBox4.Text = "FACT" Then
                            doc = "001"
                        Else
                            If Trim(ComboBox4.Text) = "" Then
                                doc = ""
                            End If

                        End If

                    End If
                End If

                If TextBox3.Text = "NIA" Then
                    ccom = "1"
                Else
                    ccom = "2"
                End If
                Dim agregar As String = "DELETE FROM custom_vianny.dbo.Qap3000 Where Empr_3a = '" + Label4.Text + "' And Ano_3a = '" + Label11.Text + "' And talm_3a = '" + TextBox1.Text + "' And ccom_3a = '" + ccom + "' And ncom_3a = '" + TextBox4.Text & TextBox5.Text + "'"
                Dim agregar1 As String = "DELETE FROM custom_vianny.dbo.Qat3000 Where Empr_3t = '" + Label4.Text + "' And Ano_3t = '" + Label11.Text + "' And talm_3t = '" + TextBox1.Text + "' And ccom_3t = '" + ccom + "' And ncom_3t = '" + TextBox4.Text & TextBox5.Text + "'"
                Dim agregar2 As String = "DELETE FROM custom_vianny.dbo.QAD3000 Where Empr_3d = '" + Label4.Text + "' And Ano_3d = '" + Label11.Text + "' And talm_3d = '" + TextBox1.Text + "' And ccom_3d = '" + ccom + "' And ncom_3d = '" + TextBox4.Text & TextBox5.Text + "'"
                ELIMINAR(agregar)
                ELIMINAR(agregar1)
                ELIMINAR(agregar2)
                Dim o As Integer
                o = DataGridView1.Rows.Count

                For i = 0 To o - 1

                    Dim cmd15 As New SqlCommand("INSERT INTO custom_vianny.dbo.Qap3000 (Empr_3a, ano_3a, talm_3a, ccom_3a, ncom_3a, item_3a, linea_3a, linea2_3a, 
				            cant_3a, unid_3a, obser1_3a, obser2_3a, obser3_3a, cantk_3a, 
				            pedido_3a, flag_3a, pieza_3a, unidk_3a, ocorte_3a, ordens_3a, Paquete_3a, OrdCos_3a)
				Values (@Empr_3a,@ano_3a,@talm_3a,@ccom_3a,@ncom_3a,@item_3a,@linea_3a, 'PT0101',@cant_3a,@unid_3a, '', '', '',@cantk_3a,@pedido_3a, 1, '',@unidk_3a,@ocorte_3a,@ordens_3a,'', '')", conx)
                    cmd15.Parameters.AddWithValue("@Empr_3a", Trim(Label4.Text))
                    cmd15.Parameters.AddWithValue("@ano_3a", Trim(Label11.Text))
                    cmd15.Parameters.AddWithValue("@talm_3a", Trim(TextBox1.Text))
                    cmd15.Parameters.AddWithValue("@ccom_3a", ccom)
                    cmd15.Parameters.AddWithValue("@ncom_3a", TextBox4.Text & TextBox5.Text)
                    cmd15.Parameters.AddWithValue("@item_3a", Trim(DataGridView1.Rows(i).Cells(0).Value))
                    cmd15.Parameters.AddWithValue("@linea_3a", Trim(DataGridView1.Rows(i).Cells(2).Value))
                    cmd15.Parameters.AddWithValue("@cantk_3a", Trim(DataGridView1.Rows(i).Cells(6).Value))
                    cmd15.Parameters.AddWithValue("@unidk_3a", Trim(DataGridView1.Rows(i).Cells(7).Value))
                    cmd15.Parameters.AddWithValue("@pedido_3a", Trim(DataGridView1.Rows(i).Cells(1).Value))
                    cmd15.Parameters.AddWithValue("@ocorte_3a", Trim(DataGridView1.Rows(i).Cells(8).Value))
                    cmd15.Parameters.AddWithValue("@ordens_3a", Trim(DataGridView1.Rows(i).Cells(9).Value))
                    cmd15.Parameters.AddWithValue("@cant_3a", Trim(DataGridView1.Rows(i).Cells(4).Value))
                    cmd15.Parameters.AddWithValue("@unid_3a", Trim(DataGridView1.Rows(i).Cells(5).Value))
                    cmd15.ExecuteNonQuery()
                Next
                If Label14.Text = 2 Then
                    Dim cmd17 As New SqlCommand("Update custom_vianny.dbo.Qag3000 SET fcom_3    = @fecha,
				                   gloa_3    = @observacion,
				                   glob_3    = '',
				                   fich_3    = @ruc,
				                   orig_3    = @origen,
				                   nombe_3   = '',
				                   adevol_3  = 0,
				                   transf_3  = 0,
				                   tdoc_3    = @doc,
				                   grm_3     = @guia,
				                   fase_3    = @area,
				                   flag_3    = 1,
				                   subfase_3 = @subfase,
				                   origen_3  = 'EXT',
				                   situa_3   = 2,
				                   Linea_3   = '',
				                   Lectura_3 = '',
				                   revisado_3 = 0
				WHERE Empr_3 = @Empr_3a And Ano_3 = @ano_3a And talm_3 =@talm_3a And ccom_3 = @ccom_3a And ncom_3 =@ncom_3a", conx)
                    cmd17.Parameters.AddWithValue("@Empr_3a", Trim(Label4.Text))
                    cmd17.Parameters.AddWithValue("@ano_3a", Trim(Label11.Text))
                    cmd17.Parameters.AddWithValue("@talm_3a", Trim(TextBox1.Text))
                    cmd17.Parameters.AddWithValue("@ccom_3a", ccom)
                    cmd17.Parameters.AddWithValue("@ncom_3a", TextBox4.Text & TextBox5.Text)
                    cmd17.Parameters.AddWithValue("@fecha", DateTimePicker1.Value.Date)
                    cmd17.Parameters.AddWithValue("@observacion", Trim(TextBox9.Text))
                    cmd17.Parameters.AddWithValue("@ruc", Trim(TextBox8.Text))
                    cmd17.Parameters.AddWithValue("@origen", Trim(TextBox7.Text))
                    cmd17.Parameters.AddWithValue("@doc", doc)
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
                    If Trim(ComboBox4.Text) = "GUIA" Then

                        cmd17.Parameters.AddWithValue("@guia", Trim(final))
                    Else

                        cmd17.Parameters.AddWithValue("@guia", Trim(TextBox12.Text))
                    End If

                    cmd17.Parameters.AddWithValue("@area", Trim(TextBox13.Text))
                    cmd17.Parameters.AddWithValue("@subfase", Trim(TextBox14.Text))
                    cmd17.ExecuteNonQuery()


                    MsgBox("LA INFORMACION DE REGISTRO CORRECTAMENTE")
                    Dim hj2 As New flog
                    Dim hj1 As New vlog
                    hj1.gmodulo = "NIA-PRODUCCION"
                    hj1.gcuser = MDIParent1.Label3.Text
                    hj1.gaccion = 2
                    hj1.gpc = My.Computer.Name
                    hj1.gfecha = DateTimePicker2.Value
                    hj1.gdetalle = Trim(TextBox1.Text) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
                    hj1.gccia = Label4.Text
                    hj2.insertar_log(hj1)
                    Dim respuesta As DialogResult

                    respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE INGRESO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then

                        Rpt_nia_nsa_prod.TextBox1.Text = TextBox1.Text
                        Rpt_nia_nsa_prod.TextBox2.Text = 1
                        Rpt_nia_nsa_prod.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
                        Rpt_nia_nsa_prod.TextBox4.Text = Label11.Text
                        Rpt_nia_nsa_prod.TextBox5.Text = Label4.Text
                        Rpt_nia_nsa_prod.Show()
                    End If
                    TextBox6.Text = ""
                    TextBox7.Text = ""
                    TextBox8.Text = ""
                    TextBox10.Text = ""
                    TextBox9.Text = ""
                    TextBox12.Text = ""
                    ComboBox4.SelectedIndex = 3
                    DataGridView1.Rows.Clear()
                    correlativo()
                Else
                    Dim cmd17 As New SqlCommand("INSERT INTO custom_vianny.dbo.Qag3000 (Empr_3, Ano_3, talm_3, ccom_3, ncom_3, fcom_3, gloa_3, glob_3, 
				            fich_3, orig_3, nombe_3, adevol_3, transf_3, tdoc_3, grm_3, fase_3, 
				            flag_3, subfase_3, origen_3, situa_3, Linea_3, Lectura_3, revisado_3)
				Values (@Empr_3a,@ano_3a,@talm_3a,@ccom_3a, @ncom_3a, @fecha, '', '',@ruc, @origen, '', 0, 0, @doc,@guia,@area, 1, @subfase, 'EXT', 2, '   ', '  ', 0)", conx)
                    cmd17.Parameters.AddWithValue("@Empr_3a", Trim(Label4.Text))
                    cmd17.Parameters.AddWithValue("@ano_3a", Trim(Label11.Text))
                    cmd17.Parameters.AddWithValue("@talm_3a", Trim(TextBox1.Text))
                    cmd17.Parameters.AddWithValue("@ccom_3a", ccom)
                    cmd17.Parameters.AddWithValue("@ncom_3a", TextBox4.Text & TextBox5.Text)
                    cmd17.Parameters.AddWithValue("@fecha", DateTimePicker1.Value.Date)
                    cmd17.Parameters.AddWithValue("@observacion", Trim(TextBox9.Text))
                    cmd17.Parameters.AddWithValue("@ruc", Trim(TextBox8.Text))
                    cmd17.Parameters.AddWithValue("@origen", Trim(TextBox7.Text))
                    cmd17.Parameters.AddWithValue("@doc", doc)
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
                    If Trim(ComboBox4.Text) = "GUIA" Then

                        cmd17.Parameters.AddWithValue("@guia", Trim(final))
                    Else

                        cmd17.Parameters.AddWithValue("@guia", Trim(TextBox12.Text))
                    End If

                    cmd17.Parameters.AddWithValue("@area", Trim(TextBox13.Text))
                    cmd17.Parameters.AddWithValue("@subfase", Trim(TextBox14.Text))
                    cmd17.ExecuteNonQuery()


                    MsgBox("LA INFORMACION DE REGISTRO CORRECTAMENTE")
                    Dim hj2 As New flog
                    Dim hj1 As New vlog
                    hj1.gmodulo = "NIA-PRODUCCION"
                    hj1.gcuser = MDIParent1.Label3.Text
                    hj1.gaccion = 1
                    hj1.gpc = My.Computer.Name
                    hj1.gfecha = DateTimePicker2.Value
                    hj1.gdetalle = Trim(TextBox1.Text) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
                    hj1.gccia = Label4.Text
                    hj2.insertar_log(hj1)
                    Dim respuesta As DialogResult

                    respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE INGRESO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then

                        Rpt_nia_nsa_prod.TextBox1.Text = TextBox1.Text
                        Rpt_nia_nsa_prod.TextBox2.Text = 1
                        Rpt_nia_nsa_prod.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
                        Rpt_nia_nsa_prod.TextBox4.Text = Label11.Text
                        Rpt_nia_nsa_prod.TextBox5.Text = Label4.Text
                        Rpt_nia_nsa_prod.Show()
                    End If
                    TextBox6.Text = ""
                    TextBox7.Text = ""
                    TextBox8.Text = ""
                    TextBox10.Text = ""
                    TextBox9.Text = ""
                    TextBox12.Text = ""

                    ComboBox4.SelectedIndex = 3
                    DateTimePicker1.Value = DateTimePicker3.Value
                    DataGridView1.Rows.Clear()
                    correlativo()


                End If
                Label14.Text = 1
            End If
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        abrir()
        Dim va As String
        Dim sql102 As String = "SELECT revisado_3 FROM custom_vianny.dbo.Qag3000 WHERE ncom_3 ='" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "' and Ano_3 ='" + Label11.Text + "' and talm_3 ='" + TextBox1.Text + "' and Empr_3 ='" + Label4.Text + "' and ccom_3 ='1'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        Dim i5 As Integer
        i5 = DataGridView1.Rows.Count
        'MsgBox(Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 1, 8))
        'MsgBox(hj)
        'MsgBox(Label2.Text)
        'MsgBox(Label3.Text)
        If Rsr2.Read() = True Then
            VA = Rsr2(0)
        End If
        Rsr2.Close()

        If va = 1 Then

            MsgBox("LA NOTA DE INGRESO ESTA REVISADA POR CONTABILIDAD NO SE PUEDE MODIFICAR")
        Else
            If RadioButton2.Checked = True Then
                MsgBox("EL DOCUMENTO ESTA ANULADO NO SE PUEDE MODIFICAR CONSULTE CON EL ADMINISTRADOR")
            Else
                Label14.Text = 2
                TextBox7.Enabled = True
                TextBox8.Enabled = True
                TextBox9.Enabled = True
                TextBox15.Enabled = True
                TextBox12.Enabled = True
                DataGridView1.Enabled = True
                Button1.Enabled = True
            End If
        End If


    End Sub
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        'Dim corre As String
        'jj.gvendedor = TextBox8.Text
        'corre = fk.buscar_correo(jj)
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='ELIMINAR_PRODUCCION '"
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
     <FONT COLOR='green'>INFORMACION DE LA NIA DE PRODUCCION</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> ANULADO <br/>
<FONT COLOR='blue'>* AREA : </FONT>" + Trim(TextBox1.Text) & "-" & Trim(TextBox2.Text) + " <br/>
<FONT COLOR='blue'>* NIA PRODUCCION : </FONT>" + Trim(TextBox4.Text) & "-" & Trim(TextBox5.Text) + " <br/>
<FONT COLOR='blue'>* USUARIO : </FONT>" + Trim(MDIParent1.Label3.Text) + " <br/>
<FONT COLOR='blue'>* PC : </FONT>" + Trim(My.Computer.Name) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(MDIParent1.Label7.Text) + " <br/>
<FONT COLOR='blue'>* TALLER : </FONT>" + Trim(TextBox10.Text) + "<br/>
</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add(Rs(0))
            Rs.Close()
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "NIA PRODUCCION N°" & Trim(TextBox1.Text) & "-" & Trim(TextBox4.Text) & Trim(TextBox5.Text)
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
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        abrir()
        Dim ccom1 As String
        If TextBox3.Text = "NIA" Then
            ccom1 = "1"
        Else
            ccom1 = "2"
        End If
        Dim cmd20 As New SqlCommand("Update custom_vianny.dbo.Qag3000 Set flag_3= 0 WHERE Empr_3 = @Empr_3a And Ano_3 = @ano_3a And talm_3 =@talm_3a And ccom_3 = @ccom_3a And ncom_3 =@ncom_3a", conx)
        cmd20.Parameters.AddWithValue("@Empr_3a", Trim(Label4.Text))
        cmd20.Parameters.AddWithValue("@ano_3a", Trim(Label11.Text))
        cmd20.Parameters.AddWithValue("@talm_3a", Trim(TextBox1.Text))
        cmd20.Parameters.AddWithValue("@ccom_3a", ccom1)
        cmd20.Parameters.AddWithValue("@ncom_3a", TextBox4.Text & TextBox5.Text)
        cmd20.ExecuteNonQuery()
        Dim cmd21 As New SqlCommand("Update custom_vianny.dbo.QaP3000 Set flag_3a= 0 WHERE Empr_3a = @Empr_3a And Ano_3a = @ano_3a And talm_3a =@talm_3a And ccom_3a = @ccom_3a And ncom_3a =@ncom_3a", conx)
        cmd21.Parameters.AddWithValue("@Empr_3a", Trim(Label4.Text))
        cmd21.Parameters.AddWithValue("@ano_3a", Trim(Label11.Text))
        cmd21.Parameters.AddWithValue("@talm_3a", Trim(TextBox1.Text))
        cmd21.Parameters.AddWithValue("@ccom_3a", ccom1)
        cmd21.Parameters.AddWithValue("@ncom_3a", TextBox4.Text & TextBox5.Text)
        cmd21.ExecuteNonQuery()
        MsgBox("SE ANULO EL COMPROBANTE CORRECTAMENTE")
        Dim hj2 As New flog
        Dim hj1 As New vlog
        hj1.gmodulo = "NIA-PRODUCCION"
        hj1.gcuser = MDIParent1.Label3.Text
        hj1.gaccion = 3
        hj1.gpc = My.Computer.Name
        hj1.gfecha = DateTimePicker2.Value
        hj1.gdetalle = Trim(TextBox1.Text) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
        hj1.gccia = Label4.Text
        hj2.insertar_log(hj1)
        ENVIAR_CORREO()
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox10.Text = ""
        TextBox9.Text = ""
        TextBox12.Text = ""
        ComboBox4.SelectedIndex = 3
        DateTimePicker1.Value = DateTimePicker3.Value
        Label14.Text = 1
        DataGridView1.Rows.Clear()
        correlativo()
    End Sub

    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()


            Dim VAL As Integer
            Select Case Trim(TextBox15.Text).Length
                Case "1" : TextBox15.Text = "01" & "0000000" & TextBox15.Text
                Case "2" : TextBox15.Text = "01" & "000000" & TextBox15.Text
                Case "3" : TextBox15.Text = "01" & "00000" & TextBox15.Text
                Case "4" : TextBox15.Text = "01" & "0000" & TextBox15.Text
                Case "5" : TextBox15.Text = "01" & "000" & TextBox15.Text
                Case "6" : TextBox15.Text = "01" & "00" & TextBox15.Text
                Case "7" : TextBox15.Text = "01" & "0" & TextBox15.Text
                Case "8" : TextBox15.Text = "01" & TextBox15.Text

            End Select


            'validar si op esta cerrafa
            Dim fim As Integer
            Dim sql103 As String = "select  finped_3 from custom_vianny.dbo.qag0300 where ncom_3 ='" + Trim(TextBox15.Text) + "' and  ccia ='" + Trim(Label4.Text) + "'"
            Dim cmd103 As New SqlCommand(sql103, conx)
            Rsr2333 = cmd103.ExecuteReader()
            If Rsr2333.Read() = True Then
                fim = Rsr2333(0)
            End If
            Rsr2333.Close()
            'fin de validar op cerrada
            If fim = 0 Then
                Dim opYaExiste As Boolean = False

                ' Recorrer todas las filas del DataGridView
                For Each row As DataGridViewRow In DataGridView1.Rows
                    ' Verificar si la columna que contiene la OP es igual a la opIngresada
                    If row.Cells("Op").Value IsNot Nothing AndAlso row.Cells("Op").Value.ToString() = TextBox15.Text.ToString().Trim() Then
                        opYaExiste = True
                        Exit For
                    End If
                Next

                ' Si la OP ya existe, mostrar un mensaje
                If opYaExiste Then
                    MessageBox.Show("La OP ingresada ya existe en el listado.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Else
                    Dim sql102 As String = "select ncom_3,linea_3,ISNULL( a.nomb_17,q.descri_3),ISNULL( a.unid_17, 'UND') from custom_vianny.dbo.qag0300 q
                left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
                where ncom_3 ='" + Trim(TextBox15.Text) + "' and q.ccia ='" + Trim(Label4.Text) + "'"
                    Dim cmd102 As New SqlCommand(sql102, conx)
                    Rsr2 = cmd102.ExecuteReader()
                    Dim i5 As Integer
                    i5 = DataGridView1.Rows.Count

                    If Rsr2.Read() = True Then
                        DataGridView1.Rows.Add()
                        DataGridView1.Rows(i5).Cells(1).Value = Rsr2(0)
                        DataGridView1.Rows(i5).Cells(2).Value = Rsr2(1)
                        DataGridView1.Rows(i5).Cells(3).Value = Rsr2(2)
                        DataGridView1.Rows(i5).Cells(7).Value = Rsr2(3)
                        DataGridView1.Rows(i5).Cells(5).Value = Rsr2(3)
                        DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                        VAL = DataGridView1.Rows(i5).Cells(0).Value
                        Select Case VAL.ToString.Length
                            Case "1" : DataGridView1.Rows(i5).Cells(0).Value = "00" & "" & VAL
                            Case "2" : DataGridView1.Rows(i5).Cells(0).Value = "0" & "" & VAL
                            Case "3" : DataGridView1.Rows(i5).Cells(0).Value = VAL
                        End Select

                        'DataGridView1.Rows(i5).Cells(4).Selected = True
                        DataGridView1.CurrentCell = DataGridView1.Rows(i5).Cells(4)
                        DataGridView1.BeginEdit(True)
                        Rsr2.Close()
                        If TextBox1.Text = "0400" Or TextBox1.Text = "0700" Then
                            abrir()
                            llenar_combo_box3()
                        End If
                        TextBox15.Text = ""

                    Else
                        MsgBox("LA OP INGRESADA NO EXISTE")
                    End If
                    Rsr2.Close()
                End If



            Else
                MsgBox("LA OP ESTA CERRADA,NO SE PUEDE HACER MOVIEMIENTO EN PRODUCCION")
            End If

        End If
    End Sub
    Sub llenar_combo_box3()
        Try

            Dim lsQuery As String = "select ocorte_3cg from custom_vianny.dbo.Qag300Cc where pedido_3cg ='" + Trim(TextBox15.Text) + "'"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            Partidaas.DataSource = loDataTable

            Partidaas.DisplayMember = "ocorte_3cg"
            Partidaas.ValueMember = "ocorte_3cg"
            Partidaas.DropDownWidth = 90
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box4(jk As String)
        Try

            Dim lsQuery1 As String = "select ocorte_3cg from custom_vianny.dbo.Qag300Cc where pedido_3cg ='" + jk + "'"
            Dim loDataTable1 As New DataTable
            Dim loDataAdapter1 As New SqlDataAdapter(lsQuery1, conx)
            loDataAdapter1.Fill(loDataTable1)
            Partidaas.DataSource = loDataTable1

            Partidaas.DisplayMember = "ocorte_3cg"
            Partidaas.ValueMember = "ocorte_3cg"
            Partidaas.DropDownWidth = 90
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.F1 Then
            det_ns_Prodc.Label2.Text = TextBox1.Text
            det_ns_Prodc.Label3.Text = 1
            det_ns_Prodc.Label5.Text = Trim(Label4.Text)
            det_ns_Prodc.Show()
        End If
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Rpt_nia_nsa_prod.TextBox1.Text = TextBox1.Text
        Rpt_nia_nsa_prod.TextBox2.Text = 1
        Rpt_nia_nsa_prod.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
        Rpt_nia_nsa_prod.TextBox4.Text = Label11.Text
        Rpt_nia_nsa_prod.TextBox5.Text = Label4.Text
        Rpt_nia_nsa_prod.Show()
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            Dim ds As New det_ns_Prodc
            ds.Label2.Text = TextBox1.Text
            ds.Label3.Text = 2
            Select Case TextBox1.Text
                Case "0300" : ds.Label4.Text = "1"
                Case "0400" : ds.Label4.Text = "2"
                Case "0600" : ds.Label4.Text = "3"
                Case "0500" : ds.Label4.Text = "4"
            End Select
            ds.Label5.Text = Trim(Label4.Text)
            ds.Label1.Text = "PROVEEDOR"
            ds.ShowDialog()
        End If
    End Sub
    Sub correlativo()
        abrir()
        Dim ccom As String

        If TextBox3.Text = "NIA" Then
            ccom = "1"
        Else
            ccom = "2"
        End If


        Dim sql1021 As String = "select top 1 ncom_3, substring(ncom_3,3,7) from custom_vianny.dbo.Qag3000 where ccom_3 = '" + ccom + "' and talm_3 ='" + TextBox1.Text + "' and Empr_3 ='" + Trim(Label4.Text) + "' and ncom_3 like '" + TextBox4.Text + "%' and Ano_3='" + Label11.Text + "' order by ncom_3 desc"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr21 = cmd1021.ExecuteReader()
        Dim i5 As Integer
        i5 = DataGridView1.Rows.Count

        If Rsr21.Read() = True Then
            TextBox5.Text = Rsr21(1) + 1
            Select Case TextBox5.Text.Length
                Case "1" : TextBox5.Text = "00000" & TextBox5.Text
                Case "2" : TextBox5.Text = "0000" & TextBox5.Text
                Case "3" : TextBox5.Text = "000" & TextBox5.Text
                Case "4" : TextBox5.Text = "00" & TextBox5.Text
                Case "5" : TextBox5.Text = "0" & TextBox5.Text
                Case "6" : TextBox5.Text = TextBox5.Text
            End Select
        Else
            TextBox5.Text = 1
            Select Case TextBox5.Text.Length
                Case "1" : TextBox5.Text = "00000" & TextBox5.Text
                Case "2" : TextBox5.Text = "0000" & TextBox5.Text
                Case "3" : TextBox5.Text = "000" & TextBox5.Text
                Case "4" : TextBox5.Text = "00" & TextBox5.Text
                Case "5" : TextBox5.Text = "0" & TextBox5.Text
                Case "6" : TextBox5.Text = TextBox5.Text
            End Select
        End If
        Rsr21.Close()
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown

        If e.KeyCode = Keys.Enter Then
            Select Case TextBox4.Text.Length
                Case "1" : TextBox4.Text = "0" & TextBox4.Text
                Case "2" : TextBox4.Text = TextBox4.Text
            End Select
            correlativo()
            TextBox5.ReadOnly = False
            TextBox5.Select()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Clear()
            abrir()
            Dim ccom As String
            If TextBox3.Text = "NIA" Then
                ccom = "1"
            Else
                ccom = "2"
            End If
            Select Case TextBox5.Text.Length
                Case "1" : TextBox5.Text = "00000" & TextBox5.Text
                Case "2" : TextBox5.Text = "0000" & TextBox5.Text
                Case "3" : TextBox5.Text = "000" & TextBox5.Text
                Case "4" : TextBox5.Text = "00" & TextBox5.Text
                Case "5" : TextBox5.Text = "0" & TextBox5.Text
                Case "6" : TextBox5.Text = TextBox5.Text
            End Select

            Dim sql1021 As String = "SELECT A.TAlm_3, 
		       		  A.CCom_3, 
		       		  A.NCom_3, 
		       		  A.FCom_3,
		       		  A.Fich_3,
		       		  A.TDoc_3,
		       		  A.Grm_3,
		       		  A.Fase_3,
		       		  A.Subfase_3,
		       		  A.ADevol_3,
		       		  A.Orig_3,
		       		  A.Gloa_3,
		       		  A.Glob_3,
		       		  A.Nombe_3,
		       		  A.Flag_3,
		       		  A.Revisado_3,
		       		  ISNULL(B.TAlm_3a, ''),
		       		  ISNULL(B.CCom_3a, ''),
		       		  ISNULL(B.NCom_3a, ''),
		       		 ISNULL( B.Linea_3a , ''),
		       		 ISNULL( B.Linea2_3a,''),
		       		 ISNULL( B.Cant_3a,0),
		       		 ISNULL( B.Unid_3a,''),
		       		 ISNULL( B.Cantk_3a,0),
		       		  ISNULL(B.Unidk_3a,''),
		       		  ISNULL(B.Pedido_3a,''),
		       		 ISNULL( B.Ordens_3a,''),
		       		 ISNULL( B.OCorte_3a,''),
		       		  ISNULL(C.Ruc_10,'') AS Ruc_10,
		       		  ISNULL(C.Nomb_10,'') AS Nomb_10,
		       		  ISNULL(D.Nomb_21e,'') AS Nomb_21e,
		       		  E.nomd_3i ,ISNULL( b.item_3a,''),isnull( c1.nomb_17,''), b.ocorte_3a
				      FROM custom_vianny.dbo.Qag3000 A LEFT JOIN custom_vianny.dbo.QAP3000  B 
				      				 ON A.empr_3 = B.empr_3a AND A.ano_3 = B.ano_3a AND A.talm_3 = B.talm_3a AND A.ccom_3 = B.ccom_3a AND A.Ncom_3 = B.NCOM_3a 
				       				 LEFT JOIN custom_vianny.dbo.CAG1000 C 
				       				 ON '01' = C.ccia AND A.fich_3 = C.fich_10 
				       				 LEFT JOIN custom_vianny.dbo.QAG210e D 
				       				 ON '01' = D.empr_21e AND a.talm_3 = d.almac_21e AND A.orig_3 = D.dest_21e 
				       				 LEFT JOIN custom_vianny.dbo.cag3i00 E 
				       				 ON A.empr_3 = E.ccia_3i AND A.tdoc_3 = E.ncom_3i
                                        LEFT JOIN custom_vianny.dbo.cag1700 C1
										 on b.Empr_3a = c1.ccia and b.linea_3a = c1.linea_17
				        Where a.empr_3 + A.ano_3 + A.talm_3  + a.ccom_3='" + Label4.Text & Label11.Text & TextBox1.Text & ccom + "'AND ( A.ncom_3 >='" + TextBox4.Text & TextBox5.Text + "' AND A.ncom_3 <= '" + TextBox4.Text & TextBox5.Text + "')
				        Order By A.Empr_3, A.Ano_3, A.talm_3, A.ccom_3, A.ncom_3"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr22 = cmd1021.ExecuteReader()

            Dim i5 As Integer
            i5 = DataGridView1.Rows.Count

            If Rsr22.Read() = True Then
                TextBox7.Text = Trim(Rsr22(10))
                TextBox6.Text = Trim(Rsr22(30))
                TextBox8.Text = Trim(Rsr22(4))
                TextBox10.Text = Trim(Rsr22(29))
                TextBox12.Text = Trim(Rsr22(6))
                TextBox9.Text = Trim(Rsr22(11))
                DateTimePicker1.Value = Trim(Rsr22(3))
                If Rsr22(5) = "009" Then
                    ComboBox4.SelectedIndex = 0
                Else
                    If Rsr22(5) = "000" Then
                        ComboBox4.SelectedIndex = 2
                    Else
                        If Rsr22(5) = "" Then
                            ComboBox4.SelectedIndex = -1
                        Else
                            ComboBox4.SelectedIndex = 1
                        End If

                    End If
                End If
                If Rsr22(14) = 1 Then
                    RadioButton1.Checked = True
                    RadioButton2.Checked = False
                    Label12.Visible = False
                Else
                    Label12.Visible = True
                    RadioButton1.Checked = False
                    RadioButton2.Checked = True
                End If
                Dim i51 As Integer
                i51 = 0
                Rsr22.Close()
                Rsr23 = cmd1021.ExecuteReader()
                While Rsr23.Read()
                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(i51).Cells(0).Value = Trim(Rsr23(32))
                    DataGridView1.Rows(i51).Cells(1).Value = Trim(Rsr23(25))
                    DataGridView1.Rows(i51).Cells(2).Value = Trim(Rsr23(19))
                    DataGridView1.Rows(i51).Cells(6).Value = Trim(Rsr23(23))
                    'DataGridView1.Rows(i51).Cells(6).Value = Rsr23(27)
                    DataGridView1.Rows(i51).Cells(9).Value = Trim(Rsr23(26))
                    DataGridView1.Rows(i51).Cells(7).Value = Trim(Rsr23(24))
                    DataGridView1.Rows(i51).Cells(3).Value = Trim(Rsr23(33))
                    DataGridView1.Rows(i51).Cells(4).Value = Trim(Rsr23(21))
                    DataGridView1.Rows(i51).Cells(5).Value = Trim(Rsr23(22))
                    If TextBox1.Text = "0400" Or TextBox1.Text = "0700" Then
                        abrir()
                        llenar_combo_box4(Rsr23(25))
                        If Trim(Rsr23(34).ToString).Length > 0 Then
                            DataGridView1.Rows(i51).Cells(6).Value = Rsr23(34).ToString
                        Else

                        End If

                    End If

                    i51 = i51 + 1
                End While
                Rsr23.Close()
                TextBox7.Select()
                TextBox7.Enabled = False
                TextBox8.Enabled = False
                TextBox9.Enabled = False
                TextBox15.Enabled = False
                TextBox12.Enabled = False
                DataGridView1.Enabled = False
                Button1.Enabled = False
            Else
                Button1.Enabled = True
                TextBox7.Enabled = True
                TextBox8.Enabled = True
                TextBox9.Enabled = True
                TextBox15.Enabled = True
                TextBox12.Enabled = True
                DataGridView1.Enabled = True
                TextBox7.Select()
            End If

        End If
    End Sub



    Private Sub TextBox5_Click(sender As Object, e As EventArgs) Handles TextBox5.Click


    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        Try
            NumConFrac(TextBox4, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        Try
            NumConFrac(TextBox5, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        Try
            NumConFrac(TextBox7, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        Try
            NumConFrac(TextBox8, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox15_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox15.KeyPress
        Try
            NumConFrac(TextBox15, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox8_Validating(sender As Object, e As CancelEventArgs) Handles TextBox8.Validating
        If Trim(DirectCast(sender, TextBox).Text).Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL RUC")
        End If
    End Sub

    Private Sub Nia_Produccion_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Me.Activate()
        Me.Focus()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        log.Label1.Text = Trim(TextBox1.Text) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
        log.Label2.Text = "NIA-PRODUCCION"
        log.Label3.Text = Label4.Text
        log.Show()
    End Sub

    Private Sub TextBox7_Validating(sender As Object, e As CancelEventArgs) Handles TextBox7.Validating
        If Trim(DirectCast(sender, TextBox).Text).Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL MOTIVO")
        End If
    End Sub

    Private Sub TextBox4_Validating(sender As Object, e As CancelEventArgs) Handles TextBox4.Validating
        If Trim(DirectCast(sender, TextBox).Text).Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL MES")
        End If
    End Sub

    Private Sub TextBox5_Validating(sender As Object, e As CancelEventArgs) Handles TextBox5.Validating
        If Trim(DirectCast(sender, TextBox).Text).Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL CORRELATIVO")
        End If
    End Sub

    Private Sub TextBox5_DoubleClick(sender As Object, e As EventArgs) Handles TextBox5.DoubleClick
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox15.Enabled = False
        TextBox12.Enabled = False
        DataGridView1.Enabled = False
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox10.Text = ""
        TextBox9.Text = ""
        TextBox12.Text = ""
        ComboBox4.SelectedIndex = 3
        DateTimePicker1.Value = DateTimePicker3.Value
        DataGridView1.Rows.Clear()
        TextBox7.Select()
        correlativo()
        Label12.Visible = False
        RadioButton1.Checked = True
        RadioButton2.Checked = False
    End Sub

    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cantidad" Then
            abrir()
            Dim va, va1, vatot, resta As Integer
            va = 0
            va1 = 0
            vatot = 0
            Dim sql1023 As String = "select ISNULL( SUM(cantk_3a),0) from custom_vianny.dbo.Qap3000 where pedido_3a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) + "' and talm_3a ='" + Trim(TextBox1.Text) + "' and Empr_3a ='" + Label4.Text + "' and ccom_3a ='2' and flag_3a ='1'"
            Dim cmd1023 As New SqlCommand(sql1023, conx)
            Rsr233 = cmd1023.ExecuteReader()
            If Rsr233.Read() = True Then
                va = Rsr233(0)
            End If
            Rsr233.Close()
            Dim sql1024 As String = "select ISNULL( SUM(cantk_3a),0) from custom_vianny.dbo.Qap3000 where pedido_3a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) + "' and talm_3a ='" + Trim(TextBox1.Text) + "' and Empr_3a ='" + Label4.Text + "' and ccom_3a ='1' and flag_3a ='1'"
            Dim cmd1024 As New SqlCommand(sql1024, conx)
            Rsr24 = cmd1024.ExecuteReader()
            If Rsr24.Read() = True Then
                va1 = Rsr24(0)
            End If
            Rsr24.Close()
            If Label14.Text = 2 Then
                vatot = 0
            Else
                vatot = va1 + Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
            End If
            resta = vatot - va1
            If vatot > va Then
                MsgBox("LA CANTIDAD INGRESADA EXCEDE EN  " + Convert.ToString(resta) + " PRENDAS PARA REGULARIZAR LA OP ")
                DataGridView1.Rows(e.RowIndex).Cells(6).Value = 0
                DataGridView1.Rows(e.RowIndex).Cells(4).Value = 0
            Else

            End If
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "A" Then
            If Trim(TextBox8.Text) <> "" Then
                Det_Os.Label4.Text = Trim(Label4.Text)
                Select Case TextBox1.Text
                    Case "0300" : Det_Os.Label5.Text = "C3"
                    Case "0400" : Det_Os.Label5.Text = "D4"
                    Case "0500" : Det_Os.Label5.Text = "E5"
                    Case "0600" : Det_Os.Label5.Text = "H8"
                    Case "0700" : Det_Os.Label5.Text = "I9"

                End Select

                Det_Os.Label6.Text = "2"
                Det_Os.Label7.Text = Trim(TextBox8.Text)
                Det_Os.Label8.Text = e.RowIndex
                Det_Os.Show()
            Else
                MsgBox("FALTA INGRESAR EL PROVEEDOR")
            End If

        End If
    End Sub



End Class