Imports System.Data.SqlClient
Public Class Actualizar_FechaDespacho
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

    Dim Rsr21, Rsr211, Rsr213 As SqlDataReader

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim cmd200 As New SqlCommand("update custom_vianny.DBO.qag0300 set frequerida_3 =@frequerida_3  where ccia =@ccia and ncom_3 =@op", conx)
        cmd200.Parameters.AddWithValue("@ccia", Label9.Text)
        cmd200.Parameters.AddWithValue("@op", Trim(TextBox2.Text))
        cmd200.Parameters.AddWithValue("@frequerida_3", Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", ""))
        cmd200.ExecuteNonQuery()
        MsgBox("SE ACTUALIZO LA FECHA DE DESPACHO DE LA OP :" + Trim(TextBox2.Text))
        limpiar()
    End Sub
    Sub limpiar()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox5.Text = ""
        TextBox3.Text = ""
        TextBox6.Text = ""
        DateTimePicker1.Value = DateTime.Now
        DateTimePicker2.Value = DateTime.Now
    End Sub
    Private Sub Actualizar_FechaDespacho_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox4.Select()
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox4.Text.Length
                Case "1" : TextBox4.Text = "010000000" & "" & TextBox4.Text
                Case "2" : TextBox4.Text = "01000000" & "" & TextBox4.Text
                Case "3" : TextBox4.Text = "0100000" & "" & TextBox4.Text
                Case "4" : TextBox4.Text = "010000" & "" & TextBox4.Text
                Case "5" : TextBox4.Text = "01000" & "" & TextBox4.Text
                Case "6" : TextBox4.Text = "0100" & "" & TextBox4.Text
                Case "7" : TextBox4.Text = "010" & "" & TextBox4.Text
                Case "8" : TextBox4.Text = "01" & "" & TextBox4.Text
                Case "9" : TextBox4.Text = "0" & "" & TextBox4.Text
                Case "10" : TextBox4.Text = TextBox4.Text
            End Select
            abrir()
            Dim sql1021 As String = "SELECT ncom_3,program_3,descri_3,frequerida_3,cants_3,cantp_3 FROM custom_vianny.DBO.qag0300 where ccia ='" + Label9.Text + "' and ncom_3 ='" + Trim(TextBox4.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr21 = cmd1021.ExecuteReader()
            Dim i5 As Integer
            i5 = 0
            If Rsr21.Read() = True Then
                TextBox1.Text = Trim(Rsr21(2))
                TextBox2.Text = Trim(Rsr21(0))
                TextBox5.Text = Trim(Rsr21(1))
                TextBox3.Text = Trim(Rsr21(4))
                TextBox6.Text = Trim(Rsr21(5))
                DateTimePicker1.Value = Rsr21(3)
            End If
            Rsr21.Close()
            TextBox4.Text = ""
        End If

    End Sub
End Class