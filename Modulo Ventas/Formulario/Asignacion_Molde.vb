Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Public Class Asignacion_Molde
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
    Dim Rsr2, Rs1, Rs19, Rs3, Rsr21, Rs25 As SqlDataReader

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim corre As String

        Dim contra As String
        If Trim(Label13.Text) = "SACOSTA" Then
            contra = "cdqpzvuxldqzonnq "
            corre = "modelista2.vianny@gmail.com"

        Else
            If Trim(Label13.Text) = "BBLAS" Then
                contra = "mlgwjildxldaaawi"
                corre = "modelista1.vianny@gmail.com"
            Else
                If Trim(Label13.Text) = "DTORRES" Then
                    contra = "wdoidyjvdkrmcqhb"
                    corre = "modelista3.vianny@gmail.com"
                End If
            End If

        End If
        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DEL MOLDE</FONT><br/><br/>
<FONT COLOR='blue'>* OP : </FONT>" + Trim(TextBox1.Text) + " <br/>
<FONT COLOR='blue'>* PO : </FONT>" + Trim(TextBox3.Text) + " <br/>
<FONT COLOR='blue'>* PRODUCTO : </FONT>" + Trim(TextBox4.Text) + " <br/>
<FONT COLOR='blue'>* TELA : </FONT>" + Trim(TextBox5.Text) + "<br/>
<FONT COLOR='blue'>* RUTA : </FONT>" + Trim(TextBox6.Text) + "<br/>
<FONT COLOR='blue'>* ESTILO : </FONT>" + Trim(TextBox7.Text) + "<br/>
<FONT COLOR='blue'>* MODELISTA : </FONT>" + Trim(TextBox9.Text) + "<br/>
<FONT COLOR='blue'>* TALLA : </FONT>" + Trim(TextBox12.Text) + "<br/>
<FONT COLOR='blue'>* LAVADO : </FONT>" + Trim(TextBox13.Text) + "<br/>
<FONT COLOR='blue'>* LAVANDERIA : </FONT>" + Trim(TextBox14.Text) + "<br/>
<FONT COLOR='blue'>* OBSERVACION : </FONT>" + Trim(TextBox11.Text) + "<br/>
<FONT COLOR='blue'>* TELA COMPLEMENTO : </FONT><br/>
" + DataGridView2.Rows(0).Cells(1).Value + "<br/>
" + DataGridView2.Rows(1).Cells(1).Value + "<br/>
" + DataGridView2.Rows(2).Cells(1).Value + "<br/>
</body>

</html>"

        Dim Mailz1 As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DEL MOLDE</FONT><br/><br/>
<FONT COLOR='blue'>* OP : </FONT>" + Trim(TextBox1.Text) + " <br/>
<FONT COLOR='blue'>* PO : </FONT>" + Trim(TextBox3.Text) + " <br/>
<FONT COLOR='blue'>* PRODUCTO : </FONT>" + Trim(TextBox4.Text) + " <br/>
<FONT COLOR='blue'>* TELA : </FONT>" + Trim(TextBox5.Text) + "<br/>
<FONT COLOR='blue'>* RUTA : </FONT>" + Trim(TextBox6.Text) + "<br/>
<FONT COLOR='blue'>* ESTILO : </FONT>" + Trim(TextBox7.Text) + "<br/>
<FONT COLOR='blue'>* MODELISTA : </FONT>" + Trim(TextBox9.Text) + "<br/>
<FONT COLOR='blue'>* TALLA : </FONT>" + Trim(TextBox12.Text) + "<br/>
<FONT COLOR='blue'>* LAVADO : </FONT>" + Trim(TextBox13.Text) + " -------  <FONT COLOR='#FF0000'>EL AREA DE UDP INFORMA QUE QUE LAVADO NO ES EL CORRECTO</FONT><br/>
<FONT COLOR='blue'>* LAVANDERIA : </FONT>" + Trim(TextBox14.Text) + "<br/>
<FONT COLOR='blue'>* OBSERVACION : </FONT>" + Trim(TextBox11.Text) + "<br/>
<FONT COLOR='blue'>* TELAS : </FONT><br/>
" + DataGridView2.Rows(0).Cells(1).Value + "<br/>
" + DataGridView2.Rows(1).Cells(1).Value + "<br/>
" + DataGridView2.Rows(2).Cells(1).Value + "<br/>
" + DataGridView2.Rows(3).Cells(1).Value + "<br/>
" + DataGridView2.Rows(4).Cells(1).Value + "<br/>
</body>

</html>"


        abrir()
        Dim co As String
        Dim sql1021011 As String = " select ISNULL(u.correo,'') FROM custom_vianny.dbo.qag0300 q left join usuario_modulo u on q.merchan_3 = u.merchan_3 WHERE ncom_3 ='" + Trim(TextBox1.Text) + "' and ccia ='" + Trim(Label9.Text) + "'"
        Dim cmd1021011 As New SqlCommand(sql1021011, conx)
        Rsr21 = cmd1021011.ExecuteReader()
        If Rsr21.Read() Then
            If Trim(Rsr21(0)) = "" Then
                co = "coordinadorit@viannysac.com"
            Else
                co = Trim(Rsr21(0))
            End If


        End If
        Rsr21.Close()
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='ASIGNACION_MOLDE    '"
        Dim cmd As New SqlCommand(sql, conx)
        Rs = cmd.ExecuteReader()
        Rs.Read()

        With message
            'Rs(0) & "," &
            .From = New System.Net.Mail.MailAddress(corre)
            .To.Add(Rs(0) & "," & corre & "," & co)
            Rs.Close()
            .IsBodyHtml = True
            If CheckBox1.Checked = True Then
                .Body = Mailz1
            Else
                .Body = Mailz
            End If

            .Subject = "Se Registro el Molde de la Op  " & TextBox1.Text
            .Priority = System.Net.Mail.MailPriority.Normal
        End With

        With smtp
            .EnableSsl = True
            .Port = "587"
            .Host = "smtp.gmail.com"
            .Credentials = New Net.NetworkCredential(corre, contra)

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
    Sub InsertarLog()
        Dim hj2 As New flog
        Dim hj1 As New vlog
        hj1.gmodulo = "MOLDE"
        hj1.gcuser = MDIParent2.Label6.Text
        hj1.gaccion = 1
        hj1.gpc = My.Computer.Name
        hj1.gfecha = DateTimePicker1.Value
        hj1.gdetalle = Trim(TextBox1.Text)
        hj1.gccia = Label9.Text
        hj2.insertar_log(hj1)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If Microsoft.VisualBasic.Len(Trim(TextBox14.Text)) = 0 Then
            MsgBox("FALTA INGRESAR LA LAVANDERIA DE DESTINO")
        Else



            Dim cmd16 As New SqlCommand("UPDATE custom_vianny.dbo.qag0300 SET mruta7_3=@RUTA, mruta8_3 = @NOMBRE, modelista_3 =@MODELISTA,scobs3_3 = @scobs3_3,scobs2_3=@scobs2_3,fcome4_3 =@fcome4_3, scobs1_3=@scobs1_3 WHERE ncom_3 =@NCOM AND ccia =@CCIA", conx)
            cmd16.Parameters.AddWithValue("@RUTA", Trim(TextBox6.Text.ToString()))
            cmd16.Parameters.AddWithValue("@NOMBRE", Trim(TextBox7.Text.ToString()))
            cmd16.Parameters.AddWithValue("@MODELISTA", Trim(TextBox8.Text.ToString()))
            cmd16.Parameters.AddWithValue("@NCOM", Trim(TextBox1.Text.ToString()))
            cmd16.Parameters.AddWithValue("@CCIA", Trim(Label9.Text))
            cmd16.Parameters.AddWithValue("@scobs3_3", Trim(TextBox11.Text.ToString()))
            cmd16.Parameters.AddWithValue("@scobs2_3", Trim(TextBox5.Text.ToString()))
            cmd16.Parameters.AddWithValue("@fcome4_3", DateTimePicker1.Value)
            cmd16.Parameters.AddWithValue("@scobs1_3", TextBox14.Text.ToString().Trim())
            If cmd16.ExecuteNonQuery() Then
                MsgBox("SE GUARDO CORRECTAMENTE LA INFORMACION")
                InsertarLog()
                If Label18.Text = "1" Then
            Pre_produccion.Button1.PerformClick()
        End If
        Label18.Text = "0"
        ENVIAR_CORREO()
        Button1.Enabled = False
        Button2.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        Else
        MsgBox("NO SE GUARDO INFORMACION")
        End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        OpenFileDialog1.ShowDialog()
        TextBox6.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button1.Enabled = False
        Button2.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox10.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub Asignacion_Molde_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            DET_FICHA_TECNICA.Close()
            DET_FICHA_TECNICA.Label3.Text = 6

            DET_FICHA_TECNICA.ShowDialog()
        Else
            If e.KeyCode = Keys.Enter Then
                DataGridView1.Rows.Clear()
                DataGridView2.Rows.Clear()
                abrir()
                Dim va As String
                Dim sql102101 As String = "SELECT ncom_3,program_3,descri_3,telaprinc_3,mruta7_3,mruta8_3,modelista_3,scobs3_3,tallador_3,ISNULL(c.nomb_10,'') AS MODELISTA,scobs2_3,alterno_3 FROM custom_vianny.dbo.qag0300 q  left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.modelista_3 = c.ruc_10 and  len(q.modelista_3) >0 WHERE ncom_3 ='" + Trim(TextBox2.Text) + "'"
                Dim cmd102101 As New SqlCommand(sql102101, conx)
                Rsr2 = cmd102101.ExecuteReader()

                If Rsr2.Read() Then
                    TextBox1.Text = Trim(Rsr2(0))
                    TextBox3.Text = Trim(Rsr2(1))
                    TextBox4.Text = Trim(Rsr2(2))
                    TextBox5.Text = Trim(Rsr2(3))
                    TextBox6.Text = Trim(Rsr2(4))
                    TextBox7.Text = Trim(Rsr2(5))
                    TextBox8.Text = Trim(Rsr2(6))
                    TextBox11.Text = Trim(Rsr2(7))
                    TextBox9.Text = Trim(Rsr2(9))
                    TextBox5.Text = Trim(Rsr2(10))
                    TextBox13.Text = Trim(Rsr2(11))
                    va = Trim(Rsr2(8))
                    Button1.Enabled = True
                    Button2.Enabled = True
                    TextBox7.Enabled = True
                    TextBox8.Enabled = True
                    Rsr2.Close()
                    abrir()
                    Dim sql22 As String = "SELECT 
		       		  ISNULL(B.Dele,'') AS NTalla1,
		       		  ISNULL(C.Dele,'') AS NTalla2,
		       		  ISNULL(D.Dele,'') AS NTalla3,
		       		  ISNULL(E.Dele,'') AS NTalla4,
		       		  ISNULL(F.Dele,'') AS NTalla5,
		       		  ISNULL(G.Dele,'') AS NTalla6,
		       		  ISNULL(H.Dele,'') AS NTalla7,
		       		  ISNULL(I.Dele,'') AS NTalla8,
		       		  ISNULL(J.Dele,'') AS NTalla9,
		       		  ISNULL(K.Dele,'') AS NTalla10
				      FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
				      				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 C
				      				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 D
				      				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 E
				      				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 F
				      				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 G
				      				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 H
				      				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 I
				      				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 J
				      				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 K
				      				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
		        Where A.CCia_tl = '" + Trim(Label9.Text) + "' AND A.Codigo_tl = '" + va + "'
		       ORDER BY  A.CCia_tl, A.Codigo_tl"
                    Dim cmd22 As New SqlCommand(sql22, conx)
                    Rs3 = cmd22.ExecuteReader()
                    Rs3.Read()
                    TextBox12.Text = Trim(Rs3(0)) + "/" + Trim(Rs3(1)) + "/" + Trim(Rs3(2)) + "/" + Trim(Rs3(3)) + "/" + Trim(Rs3(4)) + "/" + Trim(Rs3(5)) + "/" + Trim(Rs3(6)) + "/" + Trim(Rs3(7)) + "/" + Trim(Rs3(8)) + "/" + Trim(Rs3(9))
                    Rs3.Close()


                    Dim sql1021015 As String = "select tela,nomb_17 from  custom_vianny.dbo.Consumo_Op_DET p left join custom_vianny.dbo.cag1700 c on P.Ccia = c.ccia and P.tela = c.linea_17 where p.op ='" + TextBox1.Text.ToString.Trim + "' and p.ccia ='01'"
                    Dim cmd1021015 As New SqlCommand(sql1021015, conx)
                    Rs25 = cmd1021015.ExecuteReader()

                    Dim i5 As Integer
                    i5 = 0
                    DataGridView2.Rows.Add(5)
                    While Rs25.Read()

                        DataGridView2.Rows(i5).Cells(0).Value = Rs25(0)
                        DataGridView2.Rows(i5).Cells(1).Value = Rs25(1)
                        i5 = i5 + 1
                    End While
                    Rs25.Close()
                Else
                    Rsr2.Close()
                    MsgBox("LA OP NO EXISTE")
                End If


            End If
        End If
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim sql1 As String = "select nomb_10 from custom_vianny.dbo.cag1000 where ruc_10 ='" + Trim(TextBox8.Text) + "'  and tdoc_10 ='01'"
            Dim cmd1 As New SqlCommand(sql1, conx)
            Rs1 = cmd1.ExecuteReader()
            If Rs1.Read() = True Then
                TextBox9.Text = Rs1(0)
            Else
                MsgBox("SU DNI NO ESTA AGREGADO NO EXISTE")
                TextBox3.Text = ""
            End If
            Rs1.Close()
        End If

    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim sql19 As String = "select ncom_3, descri_3,mruta7_3,mruta8_3 from custom_vianny.dbo.qag0300 WHERE  modelista_3 ='" + Trim(TextBox10.Text) + "'"
            Dim cmd19 As New SqlCommand(sql19, conx)
            Rs19 = cmd19.ExecuteReader()
            Dim i5 As Integer
            i5 = 0
            While Rs19.Read()
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i5).Cells(0).Value = Rs19(0)
                DataGridView1.Rows(i5).Cells(1).Value = Rs19(1)
                DataGridView1.Rows(i5).Cells(2).Value = Rs19(2)
                DataGridView1.Rows(i5).Cells(3).Value = Rs19(3)
                i5 = i5 + 1
            End While
            Rs19.Close()
        End If
    End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.F1 Then
            det_ns_Prodc.Label2.Text = TextBox1.Text
            det_ns_Prodc.Label3.Text = 12
            det_ns_Prodc.Label4.Text = "3"
            det_ns_Prodc.Label5.Text = Trim(Label9.Text)
            det_ns_Prodc.Show()
        End If
    End Sub
End Class