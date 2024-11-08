Imports System.Data.SqlClient
Public Class EtapaMatriz
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
    Private Sub EtapaMatriz_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT DISTINCT LEFT( A.CELE,1) as codigo , LTRIM(RTRIM(CONVERT(CHAR(40), A.DELE))) as descripcion FROM custom_vianny.DBO.TAB0100 A WHERE A.CTAB='AREAMC'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 150
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
        If Label3.Text = "1" Then
            Matriz_Avios.DataGridView1.Rows(Label2.Text).Cells(1).Value = DataGridView1.Rows(e.RowIndex).Cells(1).Value
            Matriz_Avios.DataGridView1.Rows(Label2.Text).Cells(13).Value = DataGridView1.Rows(e.RowIndex).Cells(0).Value

        Else
            MatrizAviosOd.DataGridView1.Rows(Label2.Text).Cells(1).Value = DataGridView1.Rows(e.RowIndex).Cells(1).Value
            MatrizAviosOd.DataGridView1.Rows(Label2.Text).Cells(13).Value = DataGridView1.Rows(e.RowIndex).Cells(0).Value

        End If
        Me.Close()
    End Sub
End Class