Imports System.Data.SqlClient
Public Class Buscar_Partidas
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim da As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=SisGeraleyour2606")
            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Buscar_Partidas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()

        Dim cmd As New SqlDataAdapter("select distinct partidA_4 AS PARTIDA,ncom_4 AS O_TEJEDURIA,c.nomb_10 AS CLIENTE from CustomVianny.dbo.matg040f m inner join CustomVianny.dbo.cag1000 c on m.ccia_4 ='01' and m.prvtej_4 = c.fich_10 where YEAR(fcom_4) ='2023' order by partidA_4", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(1).Width = 170
        DataGridView1.Columns(2).Width = 380
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Liq_serv_Tejeduria.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Liq_serv_Tejeduria.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Me.Close()
    End Sub
End Class