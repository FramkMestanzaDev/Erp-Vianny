Imports System.Data.SqlClient
Public Class Registro_Servicios
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

    Private Sub Registro_Servicios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox3.Select()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        abrir()
        Dim cmd20 As New SqlCommand("UPDATE custom_vianny.dbo.cag1000  SET nomb_10 =@nomb_10,apepat_10=@apepat_10,apemat_10=@apemat_10,emailfe_10=@emailfe_10,fono1_10=@fono1_10 WHERE ruc_10 =@ruc_10 AND ccia ='01'", conx)
        cmd20.Parameters.AddWithValue("@nomb_10", Trim(TextBox3.Text))
        cmd20.Parameters.AddWithValue("@apepat_10", Trim(TextBox4.Text))
        cmd20.Parameters.AddWithValue("@apemat_10", Trim(TextBox5.Text))
        cmd20.Parameters.AddWithValue("@emailfe_10", Trim(TextBox6.Text))
        cmd20.Parameters.AddWithValue("@fono1_10", Trim(TextBox7.Text))
        cmd20.Parameters.AddWithValue("@ruc_10", Trim(TextBox1.Text))
        cmd20.ExecuteNonQuery()
        MsgBox("SE ACTUALIZO LA INFORMACION CORRECTAMENTE")
        SERVICIOS.ComboBox1.Focus()
        SERVICIOS.mostrar_informacion()
        Close()

    End Sub
End Class