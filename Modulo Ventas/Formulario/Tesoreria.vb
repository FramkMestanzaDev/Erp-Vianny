Imports System.Net.Mail
Public Class Tesoreria
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        BDocumentoContable.TextBox3.Text = 2
        BDocumentoContable.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox5.Text = "" Then
            MsgBox("FALTA ADJUNTAR EL ARCHIVO PDF")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox5.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox6.Text = "" Then
            MsgBox("FALTA ADJUNTAR EL ARCHIVO PDF")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox6.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox6.Text = "" Then
            MsgBox("FALTA ADJUNTAR EL ARCHIVO PDF")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox6.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox7.Text = "" Then
            MsgBox("FALTA ADJUNTAR EL ARCHIVO PDF")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox7.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient

        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("fmestanza@viannysac.com")
            .Body = "Se informa que el Area de Tesoreria Programo Fecha de Pago de La Factura" & "---" & TextBox1.Text
            .Subject = "Registro de factura"
            .Priority = System.Net.Mail.MailPriority.Normal
        End With

        With smtp
            .EnableSsl = True
            .Port = "587"
            .Host = "smtp.gmail.com"
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
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim hj As New vcontablidad
        Dim hj1 As New vcontablidad
        Dim jk As New fcontailidad

        hj.gfactura = TextBox1.Text
        hj.gperiodo = TextBox2.Text
        hj.gregistro = TextBox3.Text
        hj.gfecha = DateTimePicker1.Value
        If CheckBox1.Checked = True Then
            hj.glacrado = 1
        Else
            hj.glacrado = 0
        End If


        If jk.insertar_tesoreria(hj) = True Then
            MsgBox("SE REGISTRO LA INFORMACION CON EXITO")
            Dim hj2 As New flog
            Dim hj3 As New vlog

            hj3.gmodulo = "CONTABILIDAD-MANU"
            hj3.gcuser = MDIParent1.Label3.Text
            hj3.gaccion = 1
            hj3.gpc = My.Computer.Name
            hj3.gfecha = DateTimePicker1.Value
            hj3.gdetalle = TextBox1.Text
            hj2.insertar_log(hj3)
            ENVIAR_CORREO()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""

        Else
            MsgBox("LA INFORMACION NO SE REGISTRO")
        End If


    End Sub

    Private Sub Tesoreria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Checked = True
    End Sub
End Class