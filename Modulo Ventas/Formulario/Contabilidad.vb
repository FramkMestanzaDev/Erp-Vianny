Imports System.Net.Mail
Public Class Contabilidad
    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then

            TextBox3.ReadOnly = False
            TextBox3.ReadOnly = False
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
        End If
    End Sub

    Private Sub Contabilidad_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox3.ReadOnly = True
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        TextBox3.Text = DateTime.Now.ToString("yyyy")
        CheckBox1.Checked = True
    End Sub
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient

        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("fmestanza@viannysac.com")
            .Body = "Se informa que el Area de Contabilidad Ingreso el Registro la Factura" & "---" & TextBox1.Text
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
        Dim jk As New fcontailidad

        hj.gfactura = TextBox1.Text
        hj.gperiodo = TextBox3.Text
        hj.gregistro = TextBox4.Text
        If CheckBox1.Checked = True Then
            hj.glacrado = 1
        Else
            hj.glacrado = 0
        End If
        hj.glacrado_tesoreria = 0

        If jk.insertar(hj) = True Then
            MsgBox("SE REGISTRO LA INFORMACION CON EXITO")
            Dim hj2 As New flog
            Dim hj1 As New vlog

            hj1.gmodulo = "CONTABILIDAD-MANU"
            hj1.gcuser = MDIParent1.Label3.Text
            hj1.gaccion = 1
            hj1.gpc = My.Computer.Name
            hj1.gfecha = DateTimePicker1.Value
            hj1.gdetalle = TextBox1.Text
            hj2.insertar_log(hj1)
            ENVIAR_CORREO()
            TextBox1.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""

        Else
            MsgBox("NO SE PUDO REGISTRAR LA INFORMACION")
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
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
        If TextBox7.Text = "" Then
            MsgBox("FALTA ADJUNTAR EL ARCHIVO PDF")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox7.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox8.Text = "" Then
            MsgBox("FALTA ADJUNTAR EL ARCHIVO PDF")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox8.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox4.Text.Length < 8 Then
                MsgBox("EL FORMATO DEL REGISTRO COMPRA ES INCORRECTO, LA CANTIDAD DE DIGITOS ES ( 8 ) ", vbExclamation)
                TextBox4.Select()
            End If
        End If
    End Sub
End Class