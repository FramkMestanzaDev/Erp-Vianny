Imports System.Data.SqlClient
Public Class Articulos
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

    Dim da As New DataTable
    Private Sub Articulos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("select linea_17 AS CODIGO,nomb_17 AS DESCRIPCION, unid_17 AS UM from custom_vianny.dbo.cag1700 where ccia='02' and talm_17 ='13'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 170
        DataGridView1.Columns(1).Width = 380
    End Sub
    Dim Rsr21 As SqlDataReader
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Explosion_avios_graus.DataGridView1.Rows(Label2.Text).Cells(4).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
        Explosion_avios_graus.DataGridView1.Rows(Label2.Text).Cells(5).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        Explosion_avios_graus.DataGridView1.Rows(Label2.Text).Cells(8).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
        Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_ReporteStockFisico '02','" + Trim(Label3.Text) + "','13','" + Trim(Label3.Text) + "0101','" + Trim(Label3.Text) + "1231','" + DataGridView1.Rows(e.RowIndex).Cells(0).Value + "','" + DataGridView1.Rows(e.RowIndex).Cells(0).Value + "',NULL, 1"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr21 = cmd1021.ExecuteReader()

        If Rsr21.Read() Then
            Explosion_avios_graus.DataGridView1.Rows(Label2.Text).Cells(10).Value = Rsr21(10)
        Else
            Explosion_avios_graus.DataGridView1.Rows(Label2.Text).Cells(10).Value = 0
        End If

        Rsr21.Close()
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

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

            dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = DataGridView1.Rows.Count
                DataGridView1.Columns(0).Width = 170
                DataGridView1.Columns(1).Width = 380
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class