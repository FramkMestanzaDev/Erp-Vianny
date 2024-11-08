Imports System.Data.SqlClient
Public Class TallaMatriz
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
    Private Sub TallaMatriz_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        DataGridView1.DataSource = ""
        Dim cmd As New SqlDataAdapter("SELECT A.CORREL_16C as Codigo , CONVERT(CHAR(70) , A.Talla_16C) as Descripcion FROM custom_vianny.dbo.qag160c A WHERE A.ccia + A.ncom_16c =  '" + Label2.Text + "' + '" + Label3.Text + "'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 50
        DataGridView1.Columns(1).Width = 100
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

            dv.RowFilter = "Descripcion" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = DataGridView1.Rows.Count
                DataGridView1.Columns(0).Width = 60
                DataGridView1.Columns(1).Width = 150
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        DetalleMatrizAvios.DataGridView1.Rows(Label4.Text).Cells(1).Value = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        DetalleMatrizAvios.DataGridView1.Rows(Label4.Text).Cells(8).Value = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Me.Close()
    End Sub
End Class