Imports System.Data.SqlClient

Public Class Lista_Partidas_OT
    Public conx As SqlConnection

    Public comando As SqlCommand
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
    Private Sub Lista_Partidas_OT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()

        Dim cmd As New SqlDataAdapter("select distinct partidA_4 AS PARTIDA,ncom_4 AS O_TEJEDURIA,c.nomb_10 AS CLIENTE from custom_vianny.dbo.matg040f m inner join custom_vianny.dbo.cag1000 c on m.ccia_4 ='01' and m.prvtej_4 = c.fich_10 where YEAR(fcom_4) ='2023' order by partidA_4", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 250
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Liq_servicios_T.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Liq_servicios_T.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Me.Close()
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = DataGridView1.Rows.Count
                DataGridView1.Columns(0).Width = 100
                DataGridView1.Columns(1).Width = 250
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Liq_servicios_T.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Liq_servicios_T.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Me.Close()
    End Sub
End Class