Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class Reclamos_calidad
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
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
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim p As Integer
        p = DataGridView1.Rows.Count
        Dim agregar As String = "delete from calidad_reclamo_tela where numero = '" + Trim(TextBox1.Text) + "'"
        ELIMINAR(agregar)
        For i = 0 To p - 1
            abrir()
            Dim cmd15 As New SqlCommand("insert into calidad_reclamo_tela (procede,f_visita,hora,observaciones,cantidad,numero) values (@procede,@f_visita,@hora,@observaciones,@cantidad,@numero)", conx)
            If RadioButton1.Checked = True Then
                cmd15.Parameters.AddWithValue("@procede", 1)
            Else
                cmd15.Parameters.AddWithValue("@procede", 2)
            End If
            cmd15.Parameters.AddWithValue("@f_visita", Trim(DataGridView1.Rows(i).Cells(0).Value))
            cmd15.Parameters.AddWithValue("@hora", Trim(DataGridView1.Rows(i).Cells(1).Value))
            cmd15.Parameters.AddWithValue("@observaciones", Trim(DataGridView1.Rows(i).Cells(2).Value))
            cmd15.Parameters.AddWithValue("@cantidad", Trim(DataGridView1.Rows(i).Cells(3).Value))
            cmd15.Parameters.AddWithValue("@numero", Trim(TextBox1.Text))
            cmd15.ExecuteNonQuery()
        Next

        MsgBox("SE REGISTRO LA INFORMACION SOLICITADA")
        ENVIAR_CORREO1()
        Me.Close()
    End Sub
    Dim Rsr2 As SqlDataReader
    Private Sub Reclamos_calidad_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim sql102 As String = "SELECT * FROM calidad_reclamo_tela where numero ='" + Trim(TextBox1.Text) + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        Dim i50 As Integer
        i50 = 0
        While Rsr2.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i50).Cells(0).Value = Rsr2(2)
            DataGridView1.Rows(i50).Cells(1).Value = Rsr2(3)
            DataGridView1.Rows(i50).Cells(2).Value = Rsr2(4)
            DataGridView1.Rows(i50).Cells(3).Value = Rsr2(5)
            i50 = i50 + 1
        End While

        RadioButton1.Checked = True
    End Sub

    Private Sub Validar_Numeros(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        Dim Celda As DataGridViewCell = Me.DataGridView1.CurrentCell()

        If Celda.ColumnIndex = 3 Then

            If e.KeyChar = "."c Then

                If InStr(Celda.EditedFormattedValue.ToString, ".", CompareMethod.Text) > 0 Then

                    e.Handled = True
                Else

                    e.Handled = False
                End If
            Else

                If Len(Trim(Celda.EditedFormattedValue.ToString)) > 0 Then

                    If Char.IsNumber(e.KeyChar) Or e.KeyChar = Convert.ToChar(8) Then

                        e.Handled = False
                    Else

                        e.Handled = True
                    End If
                Else

                    If e.KeyChar = "0"c Then

                        e.Handled = True
                    Else

                        If Char.IsNumber(e.KeyChar) Or e.KeyChar = Convert.ToChar(8) Then

                            e.Handled = False
                        Else

                            e.Handled = True
                        End If
                    End If
                End If
            End If
        End If
    End Sub


    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        AddHandler e.Control.KeyPress, AddressOf Validar_Numeros
    End Sub
    Sub ENVIAR_CORREO1()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        'Dim corre As String
        'jj.gvendedor = TextBox8.Text
        'corre = fk.buscar_correo(jj)

        Dim corre As String
        jj.gvendedor = Label2.Text
        corre = fk.buscar_correo(jj)
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='HOJA_RECLAMO'"
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
     <FONT COLOR='green'>INFORMACION DE LA HOJA RECLAMO</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> VALIDADO POR CALIDAD <br/>
<FONT COLOR='blue'>* Nª RECLAMO : </FONT>" + Trim(TextBox1.Text) + " <br/>
<FONT COLOR='blue'>* MENSAJE : </FONT>" + "" + " <br/>

</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add(Rs(0) & "," & corre)
            Rs.Close()
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "HOJA RECLAMO N°" & Trim(TextBox1.Text)
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
End Class