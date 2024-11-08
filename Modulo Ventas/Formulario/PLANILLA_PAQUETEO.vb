Imports System.Data.SqlClient
Public Class PLANILLA_PAQUETEO
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
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox2.Enabled = True
            TextBox2.Select()
        Else
            TextBox2.Enabled = False
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Select Case Trim(TextBox1.Text).Length
            Case "1" : TextBox1.Text = "01" & "0000000" & TextBox1.Text
            Case "2" : TextBox1.Text = "01" & "000000" & TextBox1.Text
            Case "3" : TextBox1.Text = "01" & "00000" & TextBox1.Text
            Case "4" : TextBox1.Text = "01" & "0000" & TextBox1.Text
            Case "5" : TextBox1.Text = "01" & "000" & TextBox1.Text
            Case "6" : TextBox1.Text = "01" & "00" & TextBox1.Text
            Case "7" : TextBox1.Text = "01" & "0" & TextBox1.Text
            Case "8" : TextBox1.Text = "01" & TextBox1.Text
            Case "9" : TextBox1.Text = "0" & TextBox1.Text
        End Select
        Planilla_paqueteoo.TextBox1.Text = Label2.Text
        Planilla_paqueteoo.TextBox2.Text = TextBox1.Text
        Planilla_paqueteoo.TextBox3.Text = TextBox2.Text
        Planilla_paqueteoo.TextBox4.Text = "01"
        Planilla_paqueteoo.Show()
        Asignacion_Confeccion.TextBox1.Text = TextBox1.Text
        Asignacion_Confeccion.Show()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class