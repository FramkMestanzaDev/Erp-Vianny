Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class Orden_Produccion
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
    Private Sub Orden_Produccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim va As New vproducot
        Dim va1 As New fproductos

        va.gcia = "01"
        va.gt_pedido = "P"
        TextBox1.Text = va1.op_comercial(va)
        TextBox22.Select()
        ComboBox2.SelectedIndex = 0
        Select Case TextBox11.Text
            Case "GBEDON" : TextBox10.Text = "0013"
            Case "VINCIO" : TextBox10.Text = "0021"
            Case "DBRAVO" : TextBox10.Text = "0023"
            Case "JSALINAS" : TextBox10.Text = "0026"
            Case "GCUEVA" : TextBox10.Text = "0027"
            Case "AMENDO" : TextBox10.Text = "0024"
            Case "VGRAUS" : TextBox10.Text = "0025"
            Case "VPIZARRO" : TextBox10.Text = "0011"
            Case "JBALVIN" : TextBox10.Text = "0003"
            Case "VSILVERIO" : TextBox10.Text = "0001"
            Case "HFARJE" : TextBox10.Text = "0009"
            Case "IRODRIGUEZ" : TextBox10.Text = "0009"
            Case "EZIEGLER" : TextBox10.Text = "0009"
        End Select
        Dim func As New ftcambio
        Dim dts As New vtcambio
        dts.gfecha = DateTimePicker1.Text
        TextBox15.Text = func.mostrar_tipo_cambio(dts)
        TextBox2.Enabled = False
        TextBox9.Enabled = False
        'TextBox3.Enabled = True
        TextBox4.Enabled = False
        TextBox19.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox13.Text = "0.00"
        TextBox13.Enabled = False
        TextBox7.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        DateTimePicker1.Enabled = False
        DateTimePicker2.Enabled = False
        Button3.Enabled = False
        PictureBox2.Enabled = False
        PictureBox3.Enabled = False
        If TextBox20.Text = "SERVICIO" Then
            abrir()
            Dim nu As String
            nu = ""
            Dim Rsr19 As SqlDataReader
            Dim sql101 As String = "SELECT top 1 CONVERT (INT,SUBSTRING( program_3,7,10 )) FROM custom_vianny.dbo.qag0300 WHERE SUBSTRING( program_3,1,6 )= 'SERVIA' AND ccia ='01' order by SUBSTRING( program_3,7,10 ) desc"
            Dim cmd101 As New SqlCommand(sql101, conx)
            Rsr19 = cmd101.ExecuteReader()
            If Rsr19.Read() = False Then
                nu = 1
            Else
                nu = Rsr19(0) + 1
            End If
            Select Case nu.Length
                Case "1" : nu = "000" & "" & nu
                Case "2" : nu = "00" & "" & nu
                Case "3" : nu = "0" & "" & nu
                Case "4" : nu = nu
            End Select
            TextBox3.Text = "SERVIA" + nu
            Rsr19.Close()
        Else
            If TextBox20.Text = "TELA" Then
                abrir()
                Dim nu As String
                nu = ""
                Dim Rsr20 As SqlDataReader
                Dim sql1011 As String = "SELECT top 1 CONVERT (INT,SUBSTRING( program_3,7,10 )) FROM custom_vianny.dbo.qag0300 WHERE SUBSTRING( program_3,1,6 )= 'PROVIA' AND ccia ='01' order by SUBSTRING( program_3,7,10 ) desc"
                Dim cmd1011 As New SqlCommand(sql1011, conx)
                Rsr20 = cmd1011.ExecuteReader()
                If Rsr20.Read() = False Then
                    nu = 1
                Else
                    nu = Rsr20(0) + 1
                End If
                Select Case nu.Length
                    Case "1" : nu = "000" & "" & nu
                    Case "2" : nu = "00" & "" & nu
                    Case "3" : nu = "0" & "" & nu
                    Case "4" : nu = nu
                End Select
                TextBox3.Text = "PROVIA" + nu
                Rsr20.Close()
            Else
                abrir()
                Dim nu As String
                nu = ""
                Dim Rsr21 As SqlDataReader
                Dim sql10111 As String = "SELECT top 1 CONVERT (INT,SUBSTRING( program_3,7,10 )) FROM custom_vianny.dbo.qag0300 WHERE SUBSTRING( program_3,1,6 )= 'MUEVIA' AND ccia ='01' order by SUBSTRING( program_3,7,10 ) desc"
                Dim cmd10111 As New SqlCommand(sql10111, conx)
                Rsr21 = cmd10111.ExecuteReader()
                If Rsr21.Read() = False Then
                    nu = 1
                Else
                    nu = Rsr21(0) + 1
                End If
                Select Case nu.Length
                    Case "1" : nu = "000" & "" & nu
                    Case "2" : nu = "00" & "" & nu
                    Case "3" : nu = "0" & "" & nu
                    Case "4" : nu = nu
                End Select
                TextBox3.Text = "MUEVIA" + nu
                Rsr21.Close()
            End If
        End If
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Lista_Tela.Label2.Text = Label29.Text
        Lista_Tela.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox4.Text = "" Then
            MsgBox("FALTA INGRESAR EL CODIGO DEL PRODUCTO")
        Else
            Ficha.Label2.Text = Label33.Text
            Ficha.Show()
        End If

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)


    End Sub

    Sub ENVIAR_CORREO()
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
     <FONT COLOR='green'>INFORMACION DE LA PO </FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> CREACION <br/>
<FONT COLOR='blue'>* PO : </FONT>" + Trim(TextBox3.Text) + " <br/>
<FONT COLOR='blue'>* OP : </FONT>" + Trim(TextBox21.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(TextBox2.Text) + " <br/>
<FONT COLOR='blue'>* TIPO : </FONT>" + Trim(TextBox20.Text) + " <br/>
<FONT COLOR='blue'>* PRODUCTO : </FONT>" + Trim(TextBox12.Text) + " <br/>
<FONT COLOR='blue'>* COLOR : </FONT>" + Trim(TextBox8.Text) + "<br/>
<FONT COLOR='blue'>* CANTIDAD : </FONT>" + Trim(TextBox6.Text) + "<br/>
<FONT COLOR='blue'>* REQUERIMIENTO : </FONT>" + Trim(TextBox22.Text) + " -- " + Trim(TextBox23.Text) + "<br/>

</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("coordinadorit@viannysac.com,dbravo@viannysac.com,hrivera@viannysac.com,vgraus@viannysac.com")
            .IsBodyHtml = True
            .Body = Mailz
            If TextBox20.Text = "TELA" Then
                .Subject = "Se Genero la PO para Producir" & "    " & TextBox3.Text
            Else
                If TextBox20.Text = "MUESTRA" Then
                    .Subject = "Se Genero la PO de Muestra" & "    " & TextBox3.Text

                Else
                    .Subject = "Se Genero la PO de Servicio" & "    " & TextBox3.Text
                End If

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
     <FONT COLOR='green'>INFORMACION DE LA PO </FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> ACTUALIZACION <br/>
<FONT COLOR='blue'>* PO : </FONT>" + Trim(TextBox3.Text) + " <br/>
<FONT COLOR='blue'>* OP : </FONT>" + Trim(TextBox21.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(TextBox2.Text) + " <br/>
<FONT COLOR='blue'>* TIPO : </FONT>" + Trim(TextBox20.Text) + " <br/>
<FONT COLOR='blue'>* PRODUCTO : </FONT>" + Trim(TextBox12.Text) + " <br/>
<FONT COLOR='blue'>* COLOR : </FONT>" + Trim(TextBox8.Text) + "<br/>
<FONT COLOR='blue'>* CANTIDAD : </FONT>" + Trim(TextBox6.Text) + "<br/>
<FONT COLOR='blue'>* REQUERIMIENTO : </FONT>" + Trim(TextBox22.Text) + " -- " + Trim(TextBox23.Text) + "<br/>

</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("coordinadorit@viannysac.com,dbravo@viannysac.com,hrivera@viannysac.com,vgraus@viannysac.com")
            .IsBodyHtml = True
            .Body = Mailz
            If TextBox20.Text = "TELA" Then
                .Subject = "Se Genero la PO para Producir" & "    " & TextBox3.Text
            Else
                If TextBox20.Text = "MUESTRA" Then
                    .Subject = "Se Genero la PO de Muestra" & "    " & TextBox3.Text

                Else
                    .Subject = "Se Genero la PO de Servicio" & "    " & TextBox3.Text
                End If

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
    Private Sub limpiar()
        TextBox9.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""


        TextBox12.Text = ""
        TextBox19.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox14.Text = ""
        TextBox6.Text = ""
        TextBox13.Text = ""

        TextBox15.Text = ""
        TextBox7.Text = ""
    End Sub



    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox5.Text = "" Then
            TextBox14.Text = Mid(TextBox4.Text, 1, 7) & "C00001" & Mid(TextBox4.Text, 14, 6)
        Else
            TextBox14.Text = Mid(TextBox4.Text, 1, 7) & TextBox5.Text & Mid(TextBox4.Text, 14, 6)
        End If

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        TextBox14.Text = Mid(TextBox4.Text, 1, 7) & TextBox5.Text & Mid(TextBox4.Text, 14, 6)

    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Form_Cliente.Label1.Text = "F. PAGO"
        Form_Cliente.TextBox3.Text = 2

        Form_Cliente.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Form2.TextBox3.Text = 2
        Form2.Show()
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.F1 Then
            'Listar_Clientes.TextBox2.Text = 2
            'If TextBox10.Text = "0025" Or TextBox10.Text = "" Or TextBox10.Text = "0046" Or TextBox10.Text = "0023" Then
            '    Listar_Clientes.TextBox3.Text = 1
            'End If

            'Listar_Clientes.Show()
            Clientes.TextBox3.Text = "335"
            Clientes.Show()
        End If

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Reporte__OP.TextBox1.Text = TextBox1.Text
        Reporte__OP.TextBox2.Text = Label33.Text
        Reporte__OP.Show()
    End Sub

    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs)
        ToolTip1.SetToolTip(Button3, "GUARDAR INFORMACION")
        ToolTip1.ToolTipTitle = "ORDEN PRODUCCION"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub PictureBox3_MouseHover(sender As Object, e As EventArgs) Handles PictureBox3.MouseHover
        ToolTip1.SetToolTip(PictureBox3, "IMPRIMIR OP")
        ToolTip1.ToolTipTitle = "ORDEN PRODUCCION"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim FF As New fop
        Dim GG As New vop

        GG.gncom_3 = TextBox1.Text
        FF.actualizar_OP(GG)
        limpiar()
        Me.Close()
    End Sub

    Private Sub PictureBox2_MouseHover(sender As Object, e As EventArgs) Handles PictureBox2.MouseHover
        ToolTip1.SetToolTip(PictureBox2, "ANULAR OP")
        ToolTip1.ToolTipTitle = "ORDEN PRODUCCION"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseDown

    End Sub
    Dim pp As New DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim hj As New fop
            Dim jja As New vop
            Dim l As Integer
            jja.gncom_3 = TextBox1.Text

            pp = hj.mostarr_op(jja)

            If pp.Rows.Count <> 0 Then
                DataGridView1.DataSource = pp
                l = DataGridView1.Rows.Count
                If (Trim(DataGridView1.Rows(0).Cells(11).Value) = Trim(TextBox10.Text) Or Label30.Text = "ADMINISTRADOR" Or TextBox11.Text = "HFARJE") And (Trim(DataGridView1.Rows(0).Cells(4).Value) = "TEL" Or Trim(DataGridView1.Rows(0).Cells(4).Value) = "SER" Or Trim(DataGridView1.Rows(0).Cells(4).Value) = "MUE") Then
                    For i = 0 To l - 1
                        TextBox9.Text = DataGridView1.Rows(i).Cells(1).Value
                        TextBox2.Text = DataGridView1.Rows(i).Cells(2).Value
                        TextBox3.Text = DataGridView1.Rows(i).Cells(6).Value
                        TextBox4.Text = DataGridView1.Rows(i).Cells(5).Value

                        TextBox19.Text = DataGridView1.Rows(i).Cells(8).Value
                        TextBox5.Text = DataGridView1.Rows(i).Cells(15).Value
                        TextBox8.Text = DataGridView1.Rows(i).Cells(16).Value
                        TextBox6.Text = DataGridView1.Rows(i).Cells(9).Value
                        TextBox13.Text = DataGridView1.Rows(i).Cells(10).Value
                        TextBox7.Text = DataGridView1.Rows(i).Cells(12).Value
                        DateTimePicker1.Value = DataGridView1.Rows(i).Cells(14).Value
                        DateTimePicker2.Value = DataGridView1.Rows(i).Cells(17).Value
                        TextBox12.Text = DataGridView1.Rows(i).Cells(8).Value
                        TextBox22.Text = DataGridView1.Rows(i).Cells(19).Value
                        TextBox23.Text = DataGridView1.Rows(i).Cells(20).Value
                        TextBox24.Text = DataGridView1.Rows(i).Cells(21).Value
                        If Trim(DataGridView1.Rows(i).Cells(7).Value) = 0 Then
                            Label14.Text = "ANULADO"
                        End If
                    Next
                    TextBox18.Text = 1
                Else

                    MsgBox("LA OP NO PERTENECE A ESTE VENDEDOR O NO ES DE STOCK DE TELA")
                    TextBox2.Text = ""
                    TextBox9.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox12.Text = ""
                    TextBox19.Text = ""
                    TextBox5.Text = ""
                    TextBox8.Text = ""
                    TextBox16.Text = ""
                    TextBox14.Text = ""
                    TextBox17.Text = ""
                    TextBox6.Text = ""
                    TextBox13.Text = ""
                    TextBox7.Text = ""
                    TextBox20.Text = ""
                End If

                PictureBox3.Enabled = True
            End If
        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        TextBox2.Enabled = True
        TextBox9.Enabled = True
        TextBox3.ReadOnly = False
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox19.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox13.Enabled = True
        TextBox22.Enabled = True

        TextBox7.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
        DateTimePicker1.Enabled = True
        DateTimePicker2.Enabled = True
        Button3.Enabled = True
        PictureBox2.Enabled = True
        PictureBox3.Enabled = True
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox19.Text = TextBox12.Text
        Else
            TextBox19.Text = ""
        End If

    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged

    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Dim valor1 As String
        If TextBox1.Text = "" Or TextBox9.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox10.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or Trim(TextBox22.Text) = "" Or Trim(TextBox23.Text) = "" Then
            MsgBox("FALTAN ALGUNOS DATOS OBLIGATORIOS")
        Else
            If Label14.Text = "ANULADO" Then
                MsgBox("LA OP ESTA ANULADA NO PUEDE MODIFICARLA")
            Else
                Dim vg As New fop
                Dim vg1 As New vop
                Dim vg2 As New vop
                Dim vg3 As New vop
                Dim JJ As New vop
                Dim ko As New vop
                Dim pp As New vop
                Dim op As String
                Dim vaor As Integer
                ko.gncom_3 = TextBox1.Text
                ko.gcia = "01"
                op = vg.validar_op(ko)



                vg1.gfich_3 = TextBox9.Text
                vg1.gfcom_3 = DateTimePicker1.Value
                If ComboBox2.Text = "SOLES" Then
                    vg1.gmone_3 = "01"
                Else
                    vg1.gmone_3 = "02"
                End If
                vg1.glinea_3 = TextBox14.Text
                vg1.gFCome1_3 = DateTimePicker1.Value
                vg1.gfrequerida_3 = DateTimePicker2.Value
                vg1.gtcam_3 = TextBox15.Text
                vg1.gprogram_3 = TextBox3.Text
                vg1.gflag_3 = 1
                vg1.gdescri_3 = TextBox12.Text & "Color :" & TextBox8.Text
                vg1.galterno_3 = TextBox12.Text & "Color :" & TextBox8.Text
                vg1.gestilo_3 = "VIANNY"
                vg1.gestiloemp_3 = TextBox22.Text
                If TextBox13.Text = "" Then
                    vg1.gpfob_3 = 0
                Else
                    vg1.gpfob_3 = TextBox13.Text
                End If
                vg1.gID_REQUERIMIENTO = TextBox23.Text
                vg1.gcantp_3 = TextBox6.Text
                vg1.gcants_3 = TextBox6.Text
                vg1.gmerchan_3 = TextBox10.Text
                vg1.gbroker_3 = "0003"
                vg1.gtela1_3 = TextBox4.Text
                vg1.gtelaprinc_3 = TextBox4.Text
                vg1.gfpago_3 = "01"
                vg1.gffinprod_3 = DateTimePicker2.Value
                vg1.gobserv_3 = TextBox7.Text
                vg1.gcod_color = TextBox5.Text
                vg1.gcolor = TextBox8.Text
                If TextBox20.Text = "TELA" Then
                    vg1.gfamilia = "TEL"
                Else
                    If TextBox20.Text = "SERVICIO" Then
                        vg1.gfamilia = "SER"
                    Else
                        vg1.gfamilia = "MUE"
                    End If

                End If



                If op >= 1 Then
                    If TextBox18.Text = 1 Then
                        vg1.gncom_3 = TextBox1.Text
                        'ELIMINAR OP
                        JJ.gncom_3 = TextBox1.Text
                        JJ.gcia = "01"
                        vg.ELIMINAR_OP(JJ)
                        pp.gncom_3 = TextBox1.Text
                        pp.gcantp_3 = TextBox6.Text
                    Else
                        vaor = Convert.ToInt32(Mid(TextBox1.Text, 3, 8)) + 1

                        Select Case vaor.ToString.Length
                            Case "1" : vg1.gncom_3 = "010000000" & "" & vaor
                            Case "2" : vg1.gncom_3 = "01000000" & "" & vaor
                            Case "3" : vg1.gncom_3 = "0100000" & "" & vaor
                            Case "4" : vg1.gncom_3 = "010000" & "" & vaor
                            Case "5" : vg1.gncom_3 = "01000" & "" & vaor
                            Case "6" : vg1.gncom_3 = "0100" & "" & vaor
                            Case "7" : vg1.gncom_3 = "010" & "" & vaor
                            Case "8" : vg1.gncom_3 = "01" & "" & vaor
                        End Select
                        'ELIMINAR OP
                        Select Case vaor.ToString.Length
                            Case "1" : JJ.gncom_3 = "010000000" & "" & vaor
                            Case "2" : JJ.gncom_3 = "01000000" & "" & vaor
                            Case "3" : JJ.gncom_3 = "0100000" & "" & vaor
                            Case "4" : JJ.gncom_3 = "010000" & "" & vaor
                            Case "5" : JJ.gncom_3 = "01000" & "" & vaor
                            Case "6" : JJ.gncom_3 = "0100" & "" & vaor
                            Case "7" : JJ.gncom_3 = "010" & "" & vaor
                            Case "8" : JJ.gncom_3 = "01" & "" & vaor
                        End Select

                        JJ.gcia = "01"
                        vg.ELIMINAR_OP(JJ)
                        Select Case vaor.ToString.Length
                            Case "1" : pp.gncom_3 = "010000000" & "" & vaor
                            Case "2" : pp.gncom_3 = "01000000" & "" & vaor
                            Case "3" : pp.gncom_3 = "0100000" & "" & vaor
                            Case "4" : pp.gncom_3 = "010000" & "" & vaor
                            Case "5" : pp.gncom_3 = "01000" & "" & vaor
                            Case "6" : pp.gncom_3 = "0100" & "" & vaor
                            Case "7" : pp.gncom_3 = "010" & "" & vaor
                            Case "8" : pp.gncom_3 = "01" & "" & vaor
                        End Select
                        pp.gcantp_3 = TextBox6.Text

                        MsgBox("LA OP" & "  " & TextBox1.Text & "  " & " YA ESTA REGISTRADA-- SE ACTUALIZO A LA" & vg1.gncom_3)
                    End If

                Else
                    If op = 0 Then

                        'Eliminar op
                        JJ.gncom_3 = TextBox1.Text
                        JJ.gcia = "01"
                        vg.ELIMINAR_OP(JJ)
                        vg1.gncom_3 = TextBox1.Text
                        pp.gncom_3 = TextBox1.Text
                        pp.gcantp_3 = TextBox6.Text
                    End If



                End If
                vg2.gcia = Label33.Text
                vg2.gncom_3 = vg1.gncom_3
                vg2.gfcom_3 = DateTimePicker1.Value
                vg2.gtipo = "1"

                vg3.gncom_3 = vg1.gncom_3
                vg3.glinea_3 = TextBox4.Text
                valor1 = vg1.gncom_3
                vg.insertar_op(vg1)
                vg.insertar_cab_consumo(vg2)
                vg.insertar_cdet_consumo(vg3)
                vg.insertar_detelle_opp(pp)


                Dim Rsr20, rs21 As SqlDataReader
                Dim cant, cant1 As Integer
                Dim sql1010 As String = "select COUNT(det_req) from  det_req_comercial  where nu_requerimientio ='" + Trim(TextBox22.Text) + "'"
                Dim cmd1010 As New SqlCommand(sql1010, conx)
                Rsr20 = cmd1010.ExecuteReader()
                Rsr20.Read()
                cant = Rsr20(0)
                Rsr20.Close()

                Dim cmd15 As New SqlCommand("update det_req_comercial  set estado =1 where nu_requerimientio =@nu_requerimientio and det_req=@det", conx)
                cmd15.Parameters.AddWithValue("@nu_requerimientio", Trim(TextBox22.Text))
                cmd15.Parameters.AddWithValue("@det", Trim(TextBox23.Text))
                cmd15.ExecuteNonQuery()

                Dim sql1011 As String = "select COUNT(det_req) from  det_req_comercial  where nu_requerimientio ='" + Trim(TextBox22.Text) + "' and estado =1"
                Dim cmd1011 As New SqlCommand(sql1011, conx)
                rs21 = cmd1011.ExecuteReader()
                rs21.Read()
                cant1 = rs21(0)
                rs21.Close()
                If cant <> cant1 Then


                Else
                    Dim cmd155 As New SqlCommand("update cab_req_comercial set estado =1  where  nu_requerimientio =@nu_requerimientio", conx)
                    cmd155.Parameters.AddWithValue("@nu_requerimientio", Trim(TextBox22.Text))

                    cmd155.ExecuteNonQuery()

                End If

                MsgBox("SE REGISTRO LA ORDEN DE PRODUCCION CON EXITO")
                MsgBox("SE REGISTRO EL CONSUMO")
                TextBox21.Text = vg1.gncom_3
                Reporte__OP.TextBox1.Text = vg1.gncom_3
                Reporte__OP.Show()
                If TextBox18.Text = 1 Then
                    ENVIAR_CORREO2()
                Else
                    ENVIAR_CORREO()
                End If

                limpiar()
                Me.Close()

            End If


        End If
    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged

    End Sub

    Private Sub TextBox22_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Trim(TextBox22.Text).Length
                Case "1" : TextBox22.Text = "000000000" & "" & TextBox22.Text
                Case "2" : TextBox22.Text = "00000000" & "" & TextBox22.Text
                Case "3" : TextBox22.Text = "0000000" & "" & TextBox22.Text
                Case "4" : TextBox22.Text = "000000" & "" & TextBox22.Text
                Case "5" : TextBox22.Text = "00000" & "" & TextBox22.Text
                Case "6" : TextBox22.Text = "0000" & "" & TextBox22.Text
                Case "7" : TextBox22.Text = "000" & "" & TextBox22.Text
                Case "8" : TextBox22.Text = "00" & "" & TextBox22.Text
                Case "9" : TextBox22.Text = "0" & "" & TextBox22.Text
                Case "10" : TextBox22.Text = TextBox22.Text
            End Select

            Dim Rsr20, rs21 As SqlDataReader
            Dim cant, cant1 As Integer
            Dim sql1010 As String = "select COUNT(det_req) from  det_req_comercial  where nu_requerimientio ='" + Trim(TextBox22.Text) + "'"
            Dim cmd1010 As New SqlCommand(sql1010, conx)
            Rsr20 = cmd1010.ExecuteReader()
            Rsr20.Read()
            cant = Rsr20(0)
            Rsr20.Close()

            Dim sql1011 As String = "select COUNT(det_req) from  det_req_comercial  where nu_requerimientio ='" + Trim(TextBox22.Text) + "' and estado =1"
            Dim cmd1011 As New SqlCommand(sql1011, conx)
            rs21 = cmd1011.ExecuteReader()
            rs21.Read()
            cant1 = rs21(0)
            Rsr20.Close()
            If cant = cant1 Then
                MsgBox("EL REQUERIMIENTIO YA A SIDO ATENDIDO EN SU TOTALIDAD NO PUEDE GENERAR MAS OP")
                TextBox23.Enabled = False
            Else
                TextBox23.Select()
                TextBox23.Enabled = True
            End If
        Else
            If e.KeyCode = Keys.F1 Then
                det_reque.Label2.Text = 1
                det_reque.Show()
            End If
        End If


    End Sub

    Private Sub TextBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyCode = Keys.F1 Then
            det_reque.TextBox1.Text = TextBox22.Text
            det_reque.Label2.Text = 2
            det_reque.Show()
        End If

    End Sub

    Private Sub TextBox22_Paint(sender As Object, e As PaintEventArgs) Handles TextBox22.Paint

    End Sub

    Private Sub TextBox1_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles TextBox1.PreviewKeyDown

    End Sub
End Class