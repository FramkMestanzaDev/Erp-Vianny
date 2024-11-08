Imports System.Data.SqlClient
Public Class Detalle_Marca
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
        Dim cmd20 As New SqlCommand("update custom_vianny.dbo.TAB0100 set codigo =@codigo where  ccia + ctab='01MARCLI' and  cele =@codigo2", conx)
        cmd20.Parameters.AddWithValue("@codigo", Trim(TextBox3.Text))
        cmd20.Parameters.AddWithValue("@codigo2", Trim(TextBox1.Text))
        cmd20.ExecuteNonQuery()
        MsgBox("SE ACTUALIZO LA INFORMACION CORRECTAMENTE")
        Listado_Marca.TextBox1.Focus()
        Listado_Marca.informacion()
        Close()
    End Sub

    Private Sub Detalle_Marca_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox3.Select()
    End Sub
End Class