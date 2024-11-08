Imports System.Data.SqlClient
Public Class F_DESPACJO
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        Dim cmd15 As New SqlCommand("UPDATE custom_vianny.dbo.qag0300  SET fcome4_3=@fecha where   ncom_3 =@op AND ccia=@ccia", conx)
        cmd15.Parameters.AddWithValue("@fecha", DateTimePicker1.Value)
        cmd15.Parameters.AddWithValue("@ccia", Trim(Label3.Text))
        cmd15.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
        cmd15.ExecuteNonQuery()
        Programa_Tejeduria.mostrar23()
        Close()
    End Sub
End Class