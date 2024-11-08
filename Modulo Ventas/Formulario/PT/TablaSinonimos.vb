Imports System.Data.SqlClient
Public Class TablaSinonimos
    Public conx As SqlConnection
    Dim da As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub mostrar()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT  item_sin as Codigo, nomb_sin as Descripcion FROM custom_vianny.DBO.Cag1700_Sinonimos WHERE codigo_sin ='" + Label1.Text.ToString().Trim() + "' and ccia_sin ='" + Label2.Text.ToString().Trim() + "'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 80
        DataGridView1.Columns(1).Width = 350
    End Sub
    Private Sub TablaSinonimos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mostrar()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Guia_remision_prenda.DataGridView1.Rows(Label3.Text).Cells(2).Value = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
        Guia_remision_prenda.DataGridView1.Rows(Label3.Text).Cells(4).Value = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
        Me.Close()

    End Sub
End Class