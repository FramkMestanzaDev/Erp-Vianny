Imports System.Data.SqlClient
Public Class Registro_Cli_Comer
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

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        abrir()
        Dim cmd20 As New SqlCommand("update custom_vianny.dbo.cag1000 set sigla_10 =@sigla_10 where ccia ='01' and fich_10 =@ruc_10", conx)
        cmd20.Parameters.AddWithValue("@sigla_10", Trim(TextBox5.Text))
        cmd20.Parameters.AddWithValue("@ruc_10", Trim(TextBox1.Text))
        cmd20.ExecuteNonQuery()

        'Dim cmd21 As New SqlCommand(" update custom_vianny.dbo.cag1000 set sigla_10 =@sigla_10 where ccia ='01' and ruc_10 =@ruc_10", conx)
        'cmd21.Parameters.AddWithValue("@sigla_10", Trim(TextBox5.Text))
        'cmd21.Parameters.AddWithValue("@ruc_10", Trim(TextBox1.Text))
        'cmd21.ExecuteNonQuery()

        MsgBox("SE ACTUALIZO LA INFORMACION CORRECTAMENTE")
        Tabla_Clientes_Comerciales.TextBox1.Focus()
        Tabla_Clientes_Comerciales.mostrar_informacion()
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        Close()
    End Sub

    Private Sub Registro_Cli_Comer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox3.Select()
    End Sub
End Class