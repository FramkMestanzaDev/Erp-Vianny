Imports System.Net.Mail
Public Class Registro_Almacen


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim archivo As New OpenFileDialog
            archivo.Filter = "Archivo pdf|*.pdf"
            If archivo.ShowDialog = DialogResult.OK Then
                TextBox3.Text = archivo.FileName

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim archivo As New OpenFileDialog
            archivo.Filter = "Archivo pdf|*.pdf"
            If archivo.ShowDialog = DialogResult.OK Then
                TextBox4.Text = archivo.FileName

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim archivo As New OpenFileDialog
            archivo.Filter = "Archivo pdf|*.pdf"
            If archivo.ShowDialog = DialogResult.OK Then
                TextBox5.Text = archivo.FileName

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            Dim archivo As New OpenFileDialog
            archivo.Filter = "Archivo pdf|*.pdf"
            If archivo.ShowDialog = DialogResult.OK Then
                TextBox10.Text = archivo.FileName

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox3.Text = "" Then
            MsgBox("FALTA ADJUNTAR EL ARCHIVO PDF")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox3.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox4.Text = "" Then
            MsgBox("FALTA ADJUNTAR EL ARCHIVO PDF")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox4.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox5.Text = "" Then
            MsgBox("FALTA ADJUNTAR EL ARCHIVO PDF")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox5.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox10.Text = "" Then
            MsgBox("FALTA ADJUNTAR EL ARCHIVO PDF")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox10.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub

    Private Sub Registro_Almacen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CERRRA()
        CheckBox1.Checked = True
    End Sub
    Sub CERRRA()
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = False
        CheckBox1.Enabled = False
    End Sub
    Sub abrir()
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
        CheckBox1.Enabled = True
    End Sub
    Sub LIMPIAR()
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox2.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            TextBox2.ReadOnly = False
        End If
    End Sub


    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        If TextBox1.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Then
            MsgBox("ALGUNOS CAMPOS OBLIGATOSRIO ESTAN VACIOS")
        Else


            Dim fg As New fwilder
            Dim fg1 As New vwilder
            fg.gop = TextBox1.Text
            fg.gcotizacion = TextBox2.Text
            fg.gguia = TextBox8.Text
            fg.gdir_guia = TextBox5.Text
            fg.gnia = TextBox6.Text
            fg.gdir_nia = TextBox3.Text
            fg.gfactura = TextBox7.Text
            fg.gdir_fac = TextBox4.Text
            fg.gocompras = TextBox9.Text
            fg.gdir_ocompra = TextBox10.Text
            If CheckBox1.Checked = True Then
                fg.glacrado = 1
            Else
                fg.glacrado = 0
            End If
            fg.glacrado_contab = 0
            If fg1.insertar(fg) = True Then
                MsgBox("LA INFORMACION SE REGISTRO CON EXITO")

                Dim hj As New flog
                Dim hj1 As New vlog

                hj1.gmodulo = "ALMACEN-MANU"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = 1
                hj1.gpc = My.Computer.Name
                hj1.gfecha = DateTimePicker1.Value
                hj1.gdetalle = fg.gfactura
                hj.insertar_log(hj1)
                ENVIAR_CORREO()
                CERRRA()
                LIMPIAR()
            Else
                MsgBox("NO SE PUDO REGISTRAR LA INFORMACION")
            End If

        End If

    End Sub
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient


        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("fmestanza@viannysac.com")
            .Body = "Se informa que el Area de Almacen registro la Factura" & "---" & TextBox7.Text
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


    Private Sub TextBox7_Validated(sender As Object, e As EventArgs) Handles TextBox7.Validated

        If TextBox7.Text.Length < "13" Or Microsoft.VisualBasic.Mid(TextBox7.Text, 5, 1) <> "-" Then
            MsgBox("EL FORMATO ES INCORRECTO, LA CANTIDAD DE DIGITOS ES ( 13 ) Y ESTA SEPARADA POR ( - )", vbExclamation)
            TextBox7.Text = "F006-00000001"
            TextBox7.Select()

        End If
    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then

            If TextBox7.Text.Length < "13" Or Microsoft.VisualBasic.Mid(TextBox7.Text, 5, 1) <> "-" Then
                MsgBox("EL FORMATO ES INCORRECTO, LA CANTIDAD DE DIGITOS ES ( 13 ) Y ESTA SEPARADA POR ( - )", vbExclamation)
                TextBox7.Text = "F006-00000001"
                TextBox7.Select()

            End If
        End If
    End Sub

    Private Sub TextBox8_Validated(sender As Object, e As EventArgs) Handles TextBox8.Validated
        If TextBox8.Text.Length < "13" Or Microsoft.VisualBasic.Mid(TextBox7.Text, 5, 1) <> "-" Then
            MsgBox("EL FORMATO ES INCORRECTO, LA CANTIDAD DE DIGITOS ES ( 13 ) Y ESTA SEPARADA POR ( - )", vbExclamation)
            TextBox8.Text = "0014-00000001"
            TextBox8.Select()

        End If
    End Sub

    Private Sub TextBox8_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles TextBox8.PreviewKeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox8.Text.Length < "13" Or Microsoft.VisualBasic.Mid(TextBox7.Text, 5, 1) <> "-" Then
                MsgBox("EL FORMATO ES INCORRECTO, LA CANTIDAD DE DIGITOS ES ( 13 ) Y ESTA SEPARADA POR ( - )", vbExclamation)
                TextBox8.Text = "0014-00000001"
                TextBox8.Select()

            End If
        End If
    End Sub

    Private Sub TextBox6_Validated(sender As Object, e As EventArgs) Handles TextBox6.Validated
        If TextBox6.Text.Length < 8 Then
            MsgBox("EL FORMATO ES INCORRECTO, LA CANTIDAD DE DIGITOS ES ( 8 )", vbExclamation)
            TextBox6.Select()

        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox6.Text.Length < 8 Then
                MsgBox("EL FORMATO ES INCORRECTO, LA CANTIDAD DE DIGITOS ES ( 8 ) ", vbExclamation)

                TextBox6.Select()

            End If
        End If
    End Sub

    Private Sub TextBox9_Validated(sender As Object, e As EventArgs) Handles TextBox9.Validated
        If TextBox9.Text.Length < 8 Then
            MsgBox("EL FORMATO ES INCORRECTO, LA CANTIDAD DE DIGITOS ES ( 8 )", vbExclamation)
            TextBox9.Select()
        End If
    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox9.Text.Length < 8 Then
                MsgBox("EL FORMATO ES INCORRECTO, LA CANTIDAD DE DIGITOS ES ( 8 ) ", vbExclamation)

                TextBox9.Select()
            End If
        End If
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub
End Class