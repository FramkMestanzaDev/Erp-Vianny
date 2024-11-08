Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class planeamiento
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

    Dim ju, ju1, ju2 As New DataTable
    Private Sub planeamiento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView3.Rows.Clear()
        abrir()
        Dim Rs1, RS2, rs3, Rs111 As SqlDataReader
        Dim i5 As Integer
        Dim u As String
        Dim sql1 As String = "SELECT ncom_3, isnull(nomb_10,''),linea_3,descri_3,tallador_3,cants_3,cantp_3,porcadi_3,CONSUMO,CANT_CONSU,tela1_3,telaprinc_3 FROM custom_vianny.dbo.qag0300 Q 
        LEFT JOIN custom_vianny.dbo.cag1700 C ON Q.linea_3 = C.linea_17 AND Q.ccia = C.ccia
        LEFT JOIN custom_vianny.dbo.cag1000 G ON Q.ccia = G.ccia AND Q.fich_3 = G.fich_10
        WHERE ncom_3 ='" + TextBox1.Text + "'  and Q.ccia ='" + Trim(Label9.Text) + "' AND flag_3=1 "
        Dim cmd1 As New SqlCommand(sql1, conx)
        Rs1 = cmd1.ExecuteReader()
        Rs1.Read()
        TextBox2.Text = Rs1(1)
        TextBox3.Text = Rs1(2)
        TextBox4.Text = Rs1(3)
        TextBox5.Text = Rs1(5)
        TextBox6.Text = Rs1(7)
        TextBox7.Text = Rs1(6)

        u = Rs1(4)
        Rs1.Close()
        TextBox6.Select()
        Dim KL As New fop
        Dim OL As New vop

        OL.gcia = Label9.Text
        OL.gncom_3 = TextBox1.Text
        ju = KL.PADRON_TALLA(OL)
        ju1 = KL.PADRON_TALLA(OL)

        If ju.Rows.Count <> 0 Then
            DataGridView1.DataSource = ju

        End If
        abrir()
        Dim sql2 As String = "SELECT COUNT(cant_16d) FROM custom_vianny.dbo.Qag16dv where  ccia_16d = '" + Trim(Label9.Text) + "' And ncom_16d = '" + TextBox1.Text + "' "
        Dim cmd2 As New SqlCommand(sql2, conx)
        RS2 = cmd2.ExecuteReader()
        RS2.Read()

        If RS2(0) <> 0 Then
            ju2 = KL.PADRON_TALLA1(OL)
            DataGridView2.DataSource = ju2

        Else
            DataGridView2.DataSource = ju1
        End If
        RS2.Close()


        Dim sql11 As String = "SELECT tela,cg.nomb_17 , cxplineal,kilos,dubica,q.cantp_3, round( case when cg.familia_17 = 'TPL' then cp.cxplineal*cantp_3 else cp.kilos* q.cantp_3 end ,2)as 'K/M',cg.familia_17 FROM custom_vianny.dbo.Consumo_Op_DET  cp 
LEFT JOIN custom_vianny.dbo.cag1700 cg on cp.ccia = cg.ccia and cp.tela = cg.linea_17 
left join custom_vianny.dbo.qag0300 q on cp.ccia =q.ccia and cp.op = q.ncom_3
WHERE cp.ccia ='" + Trim(Label9.Text) + "' AND op ='" + TextBox1.Text + "' "
        Dim cmd11 As New SqlCommand(sql11, conx)
        Rs111 = cmd11.ExecuteReader()
        While Rs111.Read()
            DataGridView3.Rows.Add()
            DataGridView3.Rows(i5).Cells(0).Value = Rs111(0)
            DataGridView3.Rows(i5).Cells(1).Value = Rs111(1)
            DataGridView3.Rows(i5).Cells(2).Value = Rs111(2)
            DataGridView3.Rows(i5).Cells(3).Value = Rs111(3)
            DataGridView3.Rows(i5).Cells(4).Value = Rs111(5)
            DataGridView3.Rows(i5).Cells(5).Value = Rs111(6)
            DataGridView3.Rows(i5).Cells(6).Value = Rs111(7)
            i5 = i5 + 1
        End While
        Rs111.Close()


    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub
    Public Sub NumConFrac(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar = "." And Not CajaTexto.Text.IndexOf(".") Then
            e.Handled = True
        ElseIf e.KeyChar = "." Then
            e.Handled = False
        Else
            e.Handled = True
        End If
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
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido



        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA PROGRAMACION DE LA OP</FONT><br/><br/>
<FONT COLOR='blue'>* OP : </FONT>" + Trim(TextBox1.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(TextBox2.Text) + " <br/>
<FONT COLOR='blue'>* PRODUCTO : </FONT>" + Trim(TextBox4.Text) + " <br/>
<FONT COLOR='blue'>* CANTIDAD SOLICITADA : </FONT>" + TextBox5.Text + "<br/>
<FONT COLOR='blue'>* ADICIONAL : </FONT>" + TextBox6.Text + "<br/>
<FONT COLOR='blue'>* CANTIDAD PROGRAMADA : </FONT>" + TextBox7.Text + "<br/>
</body>
</html>"
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='PRODUCCION          '"
        Dim cmd As New SqlCommand(sql, conx)
        Rs = cmd.ExecuteReader()
        Rs.Read()

        With message

            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add(Rs(0))
            Rs.Close()
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "Se Registro La Programacion de la Op  " & TextBox1.Text
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
            MessageBox.Show("EL mensaje de correo ha sido enviado.", "Correo enviado",
                             MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error al enviar correo",
                             MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        Dim pl As Integer
        pl = DataGridView3.Rows.Count

        If Trim(Label9.Text) = "01" Then
            'If pl <= 1 Or Trim(TextBox7.Text) = "0.00" Then
            '    MsgBox("NO SE INGRESO CONSUMO DE ESTA OP EN EL SISTEMA - REVISAR CON UDP")

            'Else



            Dim RS4, RS45 As SqlDataReader
            Dim J, j1, estado As String
            Dim agregar As String = "delete from custom_vianny.dbo.Qag16dv where  ccia_16d = '" + Trim(Label9.Text) + "' And ncom_16d = '" + Trim(TextBox1.Text) + "'"
                ELIMINAR(agregar)
                Dim cmd15 As New SqlCommand("Update custom_vianny.dbo.Qag0300 SET cantp_3 = @cantp, porcadi_3 = @porc Where ccia = @ccia  And ncom_3 = @op", conx)
                cmd15.Parameters.AddWithValue("@cantp", Trim(TextBox7.Text))
                cmd15.Parameters.AddWithValue("@porc", Trim(TextBox6.Text))
                cmd15.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
                cmd15.Parameters.AddWithValue("@ccia", Trim(Label9.Text))
                cmd15.ExecuteNonQuery()
                Dim oj As Integer
                oj = DataGridView2.Columns.Count
                For i = 0 To oj - 1

                    Dim sql2 As String = "SELECT correl_16c FROM custom_vianny.dbo.qag160c where  ccia = '" + Trim(Label9.Text) + "' And ncom_16c = '" + Trim(TextBox1.Text) + "' and talla_16c ='" + Me.DataGridView2.Columns.Item(i).Name.ToString + "'"
                    Dim cmd2 As New SqlCommand(sql2, conx)
                    RS4 = cmd2.ExecuteReader()
                    RS4.Read()
                    J = RS4(0)
                    RS4.Close()
                    Dim cmd20 As New SqlCommand("INSERT INTO custom_vianny.dbo.Qag16dv (ccia_16d, ncom_16d, llave_16d, color_16d, correl_16d, cant_16d) Values (@ccia, @op, 0, '01', @CORREL, @CANTIDAD)", conx)
                    cmd20.Parameters.AddWithValue("@ccia", Trim(Label9.Text))
                    cmd20.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
                    cmd20.Parameters.AddWithValue("@CANTIDAD", Trim(DataGridView2.Rows(0).Cells(i).Value))
                    cmd20.Parameters.AddWithValue("@CORREL", J)
                    cmd20.ExecuteNonQuery()

                Next

            If Convert.ToDouble(TextBox7.Text) > 0 Then
                Dim sql3 As String = " select ISNULL( ncorte_3,0) from  custom_vianny.DBO.qag0300  WHERE ncom_3 =" + Trim(TextBox1.Text) + " And ccia ='" + Trim(Label9.Text) + "'"
                Dim cmd3 As New SqlCommand(sql3, conx)
                RS45 = cmd3.ExecuteReader()
                If RS45.Read() Then
                    j1 = RS45(0)
                Else
                    j1 = 0
                End If

                RS45.Close()
                If j1 = "2" Then
                    estado = "2"
                Else
                    estado = "1"
                End If
                Dim cmd16 As New SqlCommand("UPDATE custom_vianny.DBO.qag0300 Set ncorte_3='" + estado + "' WHERE ncom_3 =@OP And ccia =@CCIA", conx)
                cmd16.Parameters.AddWithValue("@OP", Trim(TextBox1.Text))
                cmd16.Parameters.AddWithValue("@CCIA", Trim(Label9.Text))
                cmd16.ExecuteNonQuery()
                ENVIAR_CORREO()
            End If




            'End If

            MsgBox("SE GUARDO LA INFORMACION SOLICITADA")

            Me.Close()
        Else
            Dim RS4 As SqlDataReader
            Dim J As String
            Dim agregar As String = "delete from custom_vianny.dbo.Qag16dv where  ccia_16d = '" + Trim(Label9.Text) + "' And ncom_16d = '" + Trim(TextBox1.Text) + "'"
            ELIMINAR(agregar)
                    Dim cmd15 As New SqlCommand("Update custom_vianny.dbo.Qag0300 SET cantp_3 = @cantp, porcadi_3 = @porc Where ccia = @ccia  And ncom_3 = @op", conx)
                    cmd15.Parameters.AddWithValue("@cantp", Trim(TextBox7.Text))
                    cmd15.Parameters.AddWithValue("@porc", Trim(TextBox6.Text))
                    cmd15.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
                    cmd15.Parameters.AddWithValue("@ccia", Trim(Label9.Text))
                    cmd15.ExecuteNonQuery()
                    Dim oj As Integer
                    oj = DataGridView2.Columns.Count
                    For i = 0 To oj - 1

                        Dim sql2 As String = "SELECT correl_16c FROM custom_vianny.dbo.qag160c where  ccia = '" + Trim(Label9.Text) + "' And ncom_16c = '" + Trim(TextBox1.Text) + "' and talla_16c ='" + Me.DataGridView2.Columns.Item(i).Name.ToString + "'"
                        Dim cmd2 As New SqlCommand(sql2, conx)
                        RS4 = cmd2.ExecuteReader()
                        RS4.Read()
                        J = RS4(0)
                        RS4.Close()
                        Dim cmd20 As New SqlCommand("INSERT INTO custom_vianny.dbo.Qag16dv (ccia_16d, ncom_16d, llave_16d, color_16d, correl_16d, cant_16d) Values (@ccia, @op, 0, '01', @CORREL, @CANTIDAD)", conx)
                        cmd20.Parameters.AddWithValue("@ccia", Trim(Label9.Text))
                        cmd20.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
                        cmd20.Parameters.AddWithValue("@CANTIDAD", Trim(DataGridView2.Rows(0).Cells(i).Value))
                        cmd20.Parameters.AddWithValue("@CORREL", J)
                        cmd20.ExecuteNonQuery()

                    Next
                    MsgBox("SE GUARDO LA INFORMACION SOLICITADA")
                    Me.Close()
                End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        NumConFrac(TextBox6, e)
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim u, u1 As Integer
            Dim consumo As Double
            u1 = DataGridView3.Rows.Count
            If TextBox5.Text = 0 Or TextBox5.Text = "" Then
                MsgBox("LA CANTIDAD SOLICITADA ESTA VACIA, DEBE TENER UN VALOR")
            Else
                If Trim(Label9.Text) = "01" Then
                    If u1 = 0 Then
                        MsgBox("NO EXISTE CONSUMO INGRESADO PARA ESTA OP, NO SE PUEDE AGREGAR EL EXCENDENTE DE PRODUCCION")
                    Else
                        TextBox7.Text = Math.Round(TextBox5.Text + ((TextBox5.Text * TextBox6.Text) / 100))
                        Dim pl, cat As Double
                        cat = Math.Round(TextBox5.Text + ((TextBox5.Text * TextBox6.Text) / 100))

                        u = DataGridView2.Columns.Count

                        For i = 0 To u - 1
                            pl = DataGridView1.Rows(0).Cells(i).Value + ((DataGridView1.Rows(0).Cells(i).Value * TextBox6.Text) / 100)
                            DataGridView2.Rows(0).Cells(i).Value = Math.Round(pl)

                        Next


                        For i1 = 0 To u1 - 1

                            DataGridView3.Rows(i1).Cells(4).Value = cat
                            If Trim(DataGridView3.Rows(i1).Cells(6).Value) = "TPL" Then
                                consumo = Convert.ToDouble(DataGridView3.Rows(i1).Cells(4).Value) * Convert.ToDouble(DataGridView3.Rows(i1).Cells(2).Value)
                                DataGridView3.Rows(i1).Cells(5).Value = Math.Round(consumo, 2)
                            Else
                                consumo = Convert.ToDouble(DataGridView3.Rows(i1).Cells(4).Value) * Convert.ToDouble(DataGridView3.Rows(i1).Cells(3).Value)
                                DataGridView3.Rows(i1).Cells(5).Value = Math.Round(consumo, 2)
                            End If
                        Next
                    End If
                Else
                    TextBox7.Text = Math.Round(TextBox5.Text + ((TextBox5.Text * TextBox6.Text) / 100))
                    Dim pl, cat As Double
                    Dim cant As Integer
                    cat = Math.Round(TextBox5.Text + ((TextBox5.Text * TextBox6.Text) / 100))

                    u = DataGridView2.Columns.Count

                    For i = 0 To u - 1
                        pl = DataGridView1.Rows(0).Cells(i).Value + ((DataGridView1.Rows(0).Cells(i).Value * TextBox6.Text) / 100)
                        DataGridView2.Rows(0).Cells(i).Value = Math.Round(pl)
                        cant = cant + Convert.ToInt32(DataGridView2.Rows(0).Cells(i).Value)
                    Next

                    If Convert.ToInt32(TextBox7.Text) = cant Then
                        TextBox7.Text = Convert.ToInt32(TextBox7.Text)
                    Else
                        TextBox7.Text = cant

                    End If
                End If


                End If

        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.F1 Then
            Productos_Tela.Label2.Text = Label9.Text
            Productos_Tela.Show()
        End If
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs)
        'If e.KeyCode = Keys.Enter Then
        '    If TextBox5.Text = 0 Or TextBox5.Text = "" Then
        '        MsgBox("LA CANTIDAD SOLICITADA ESTA VACIA, DEBE TENER UN VALOR")
        '    Else
        '        TextBox11.Text = TextBox7.Text * TextBox10.Text

        '    End If

        'End If
    End Sub

    Private Sub DataGridView2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellEndEdit
        Dim y, p, porc As Integer
        Dim op As Double
        y = DataGridView2.Columns.Count
        p = 0
        Do While (p < y)
            If e.ColumnIndex = p Then
                For i = 0 To y - 1
                    op = op + DataGridView2.Rows(0).Cells(i).Value
                Next
                TextBox7.Text = op
                porc = Math.Round((TextBox7.Text * 100) / TextBox5.Text)
                TextBox6.Text = porc - 100
            End If
            p = p + 1
        Loop




    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

    End Sub
End Class