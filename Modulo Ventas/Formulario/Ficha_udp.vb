Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Public Class Ficha_udp
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand
    Dim Rsr2 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            DET_FICHA_TECNICA.Close()
            DET_FICHA_TECNICA.Label3.Text = 7

            DET_FICHA_TECNICA.ShowDialog()
        Else
            If e.KeyCode = Keys.Enter Then
                Select Case Trim(TextBox2.Text).Length
                    Case "1" : TextBox2.Text = "01" & "0000000" & TextBox2.Text
                    Case "2" : TextBox2.Text = "01" & "000000" & TextBox2.Text
                    Case "3" : TextBox2.Text = "01" & "00000" & TextBox2.Text
                    Case "4" : TextBox2.Text = "01" & "0000" & TextBox2.Text
                    Case "5" : TextBox2.Text = "01" & "000" & TextBox2.Text
                    Case "6" : TextBox2.Text = "01" & "00" & TextBox2.Text
                    Case "7" : TextBox2.Text = "01" & "0" & TextBox2.Text
                    Case "8" : TextBox2.Text = "01" & TextBox2.Text
                    Case "9" : TextBox2.Text = "0" & TextBox2.Text
                End Select
                abrir()
                Dim sql102 As String = "select F.*,c.nomb_10 from Ficha_diseno F LEFT JOIN custom_vianny.DBO.cag1000 c on f.trabajador =c.ruc_10 where op ='" + Trim(TextBox2.Text) + "'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()
                Dim i5 As Integer
                i5 = 0

                Rsr2.Read()
                TextBox1.Text = Rsr2(1)
                TextBox3.Text = Trim(Rsr2(2))
                TextBox4.Text = Trim(Rsr2(3))
                TextBox5.Text = Trim(Rsr2(4))
                TextBox6.Text = Trim(Rsr2(5))
                TextBox7.Text = Trim(Rsr2(11))
                TextBox9.Text = Trim(Rsr2(12))
                TextBox8.Text = Rsr2(6)
                TextBox11.Text = Trim(Rsr2(10))
                If Trim(Rsr2(7)) = "1" Then
                    CheckBox2.Checked = True
                Else
                    CheckBox2.Checked = False
                End If
                If Trim(Rsr2(8)) = "1" Then
                    CheckBox3.Checked = True
                Else
                    CheckBox3.Checked = False
                End If
                If Trim(Rsr2(9)) = "1" Then
                    CheckBox4.Checked = True
                Else
                    CheckBox4.Checked = False
                End If
                Rsr2.Close()
                TextBox2.Text = ""
                Button3.Enabled = True
            End If

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button1.Enabled = False
        Button2.Enabled = False
        Button5.Enabled = False
        TextBox8.Enabled = False
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""

        TextBox9.Text = ""
        TextBox11.Text = ""
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        TextBox2.Select()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        TextBox6.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub
    Dim Rs1 As SqlDataReader
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
    Sub ENVIAR_CORREO2()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim corre As String
        jj.gvendedor = TextBox8.Text
        corre = fk.buscar_correo(jj)
        Dim fco, faca, fav, ESTADO As String
        If Label9.Text = 1 Then
            ESTADO = "MODIFICO"
        Else
            ESTADO = "CREO"
        End If
        If CheckBox2.Checked = True Then
            fco = "SI"
        Else
            fco = "NO"
        End If
        If CheckBox3.Checked = True Then
            fav = "SI"
        Else
            fav = "NO"
        End If
        If CheckBox4.Checked = True Then
            faca = "SI"
        Else
            faca = "NO"
        End If

        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
         <FONT COLOR='green'>INFORMACION DE LA FICHA TECNICA DISEÑO</FONT><br/><br/>

    <FONT COLOR='blue'>* ESTADO : </FONT> " + ESTADO + " <br/>
    <FONT COLOR='blue'>* OP : </FONT>" + Trim(TextBox1.Text) + " <br/>
    <FONT COLOR='blue'>* PO : </FONT>" + Trim(TextBox3.Text) + " <br/>
    <FONT COLOR='blue'>* MODELO : </FONT>" + Trim(TextBox4.Text) + " <br/>
    <FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(TextBox5.Text) + " <br/>
    <FONT COLOR='blue'>* RUTA FICHA : </FONT>" + Trim(TextBox6.Text) + "<br/>
    <FONT COLOR='blue'>* RUTA HOJA MEDIDA : </FONT>" + Trim(TextBox7.Text) + "<br/>
    <FONT COLOR='blue'>* TRABAJADOR : </FONT>" + Trim(TextBox9.Text) + "<br/>
    <FONT COLOR='blue'>* OBSERVACION : </FONT>" + Trim(TextBox11.Text) + "<br/>
    <FONT COLOR='blue'>* FICHA COSTURA : </FONT>" + Trim(fco) + "<br/>
    <FONT COLOR='blue'>* FICHA AVIOS : </FONT>" + Trim(fav) + "<br/>
    <FONT COLOR='blue'>* FICHA ACABADOS : </FONT>" + Trim(faca) + "<br/>
    </body>

    </html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("mvara@viannysac.com,epalomino@viannysac.com,jefemanufactura@viannysac.com,otarazon@viannysac.com,jcespedes@viannysac.com,mperalta@viannysac.com,mhuaman@viannysac.com,ecallata@viannysac.com,serviciosexternos@viannysac.com,asilva@viannysac.com,bblas@viannysac.com,gbalvin@viannysac.com,vsilverio@viannysac.com,wrosales@viannysac.com")
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "Ficha Tecnica Diseño - OP N°" & TextBox1.Text
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

        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()
        Dim agregar As String = "delete  from Ficha_diseno where op ='" + Trim(TextBox1.Text) + "'"
        ELIMINAR(agregar)
        Dim cmd16 As New SqlCommand("insert into Ficha_diseno (ccia,op,po,modelo,cliente,ruta,trabajador,f_costuta,f_avios,f_acabados,observacion,ruta_medida) values (@ccia,@op,@po,@modelo,@cliente,@ruta,@trabajador,@f_costuta,@f_avios,@f_acabados,@observacion,@ruta_medida)", conx)
        cmd16.Parameters.AddWithValue("@ccia", Trim(Label10.Text))
        cmd16.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
        cmd16.Parameters.AddWithValue("@po", Trim(TextBox3.Text))
        cmd16.Parameters.AddWithValue("@modelo", Trim(TextBox4.Text))
        cmd16.Parameters.AddWithValue("@cliente", Trim(TextBox5.Text))
        cmd16.Parameters.AddWithValue("@ruta", Trim(TextBox6.Text))
        cmd16.Parameters.AddWithValue("@ruta_medida", Trim(TextBox7.Text))
        cmd16.Parameters.AddWithValue("@trabajador", Trim(TextBox8.Text))
        If CheckBox2.Checked = True Then
            cmd16.Parameters.AddWithValue("@f_costuta", "1")
        Else
            cmd16.Parameters.AddWithValue("@f_costuta", "0")
        End If
        If CheckBox3.Checked = True Then
            cmd16.Parameters.AddWithValue("@f_avios", "1")
        Else
            cmd16.Parameters.AddWithValue("@f_avios", "0")

        End If
        If CheckBox4.Checked = True Then
            cmd16.Parameters.AddWithValue("@f_acabados", "1")
        Else
            cmd16.Parameters.AddWithValue("@f_acabados", "0")

        End If

        cmd16.Parameters.AddWithValue("@observacion", Trim(TextBox11.Text))
        If cmd16.ExecuteNonQuery() Then
            MsgBox("SE GUARDO CORRECTAMENTE LA INFORMACION")
            Dim hj2 As New flog
            Dim hj11 As New vlog
            hj11.gmodulo = "FICHA_UDP"
            hj11.gcuser = MDIParent1.Label3.Text
            If Label9.Text = 0 Then
                hj11.gaccion = "CREAR"
            Else
                If Label9.Text = 1 Then
                    hj11.gaccion = "MODIFICACION"
                End If
            End If

            hj11.gpc = My.Computer.Name
            hj11.gfecha = String.Format("{0:G}", DateTime.Now)
            hj11.gdetalle = Trim(TextBox1.Text)
            hj11.gccia = Label10.Text
            hj2.insertar_log(hj11)
            ENVIAR_CORREO2()

            Button1.Enabled = False
            Button2.Enabled = False
            Button5.Enabled = False
            TextBox8.Enabled = False
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox8.Text = ""
            TextBox7.Text = ""
            TextBox9.Text = ""
            TextBox11.Text = ""
            CheckBox2.Checked = False
            CheckBox3.Checked = False
            CheckBox4.Checked = False
        Else
            MsgBox("NO SE GUARDO INFORMACION")
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label9.Text = 1
        Button1.Enabled = True
        Button5.Enabled = True
        Button4.Enabled = True
        Button2.Enabled = True
        TextBox11.Enabled = True
        TextBox8.Enabled = True

        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
    End Sub

    Private Sub Ficha_udp_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        OpenFileDialog1.ShowDialog()
        TextBox7.Text = OpenFileDialog1.FileName
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
End Class