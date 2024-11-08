Imports System.Data.SqlClient
Public Class Detalle_Asignacion
    Public conx As SqlConnection
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim da As New DataTable
    Private Sub Detalle_Asignacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("select op as OP, CORTE AS CORTE,RUC AS RUC,PROVEEDOR AS CLIENTE from asignacion_cos_aca WHERE AREA ='" + Label2.Text + "'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 50
        DataGridView1.Columns(2).Width = 120
        DataGridView1.Columns(1).Width = 300
    End Sub
End Class