Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class EnviarCorreo
    Public conx As SqlConnection
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub ENVIAR_CORREO()
        abrir()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim cabecera, usuario, contras As String
        usuario = ""
        contras = ""
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "SELECT correo,contrasena FROM usuario_modulo where usuario ='" + Trim(Label3.Text.ToString().Trim()) + "'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        If Rsr1991.Read() Then
            usuario = Rsr1991(0).ToString().Trim()
            contras = Rsr1991(1).ToString().Trim()
        End If
        Rsr1991.Close()
        If Label5.Text.ToString().Trim() = 0 Then
            cabecera = "INFORMACION DEL CONSUMO DE TELA"
        Else
            cabecera = "INFORMACION DEL REQUERIMIENTO DE AVIOS OP " + Label4.Text.ToString().Trim()
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
                <th align='center'><font color='blue'>CABECERA</font></th>
                <th align='center'><font color='blue'>DETALLE</font></th>
                <th align='center'><font color='blue'>FECHA</font></th>
                
            </tr>
        </thead>
        <tbody>
            <tr>              
                <td  align='center'> " + TextBox1.Text.ToString().Trim() + "</td>                       
                <td  align='center'>" + TextBox2.Text.ToString().Trim() + " </td>    
                <td  align='center'>" + DateTimePicker1.Value + " </td>    
            </tr>
        </tbody>
    </table>
</body>
</html>"
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='PRODUCCION          '"
        Dim cmd As New SqlCommand(sql, conx)
        Rs = cmd.ExecuteReader()
        Rs.Read()

        With message
            MsgBox(Rs(0).ToString().Trim())
            .From = New System.Net.Mail.MailAddress(usuario)
            .To.Add(Rs(0).ToString().Trim())
            .To.Add(usuario)
            Rs.Close()
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = cabecera
            .Priority = System.Net.Mail.MailPriority.Normal
        End With


        With smtp

            .EnableSsl = True
            .Port = "587"
            If Label3.Text.ToString().Trim() = "JMEDINA" Then
                .Host = "smtp.gmail.com"
            Else
                .Host = "smtppro.zoho.com"
            End If

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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If validar() Then
            ENVIAR_CORREO()
            TextBox1.Text = ""
            TextBox2.Text = ""
            Label3.Text = ""
            Label4.Text = ""
            Label5.Text = "0"
            Me.Close()
        End If

    End Sub
    Private Function validar() As Boolean
        If TextBox1.Text.ToString().Trim().Length = 0 Then
            MsgBox("FALTA INGRESAR INFORMACION EN LA CABECERA")
            Return False
        End If
        If TextBox2.Text.ToString().Trim().Length = 0 Then
            MsgBox("FALTA INGRESAR INFORMACION EN EL DETALLE")
            Return False
        End If
        Return True
    End Function

    Private Sub EnviarCorreo_Load(sender As Object, e As EventArgs) Handles MyBase.Load



    End Sub
End Class