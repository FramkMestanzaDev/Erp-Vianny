Imports System.Data.SqlClient
Public Class DET_FICHA_CONSUMO
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
    Private Sub DET_FICHA_CONSUMO_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        da.Clear()
        Dim cmd1 As New SqlDataAdapter("select ncom_3 AS OP, ISNULL( c.nomb_10,'') AS CLIENTE,descri_3 AS PRODUCTO,estilo_3 AS ESTILO from custom_vianny.dbo.qag0300 q left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 = c.fich_10 where q.ccia ='01' and year(fcom_3) =year(GETDATE()) and ncom_3 like '01%' and  (merchan_3 <> '0009' or  merchan_3 <> '0026') and  (broker_3 ='0001' or broker_3 ='0000') order by ncom_3 desc", conx)
        cmd1.Fill(da)
        If da.Rows.Count <> 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(0).Width = 100
            DataGridView1.Columns(1).Width = 180
            DataGridView1.Columns(2).Width = 200
            DataGridView1.Columns(3).Width = 170
        Else
            DataGridView1.DataSource = Nothing
        End If


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex = -1 Then

        Else
            FICHA_CONSUMO.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)

            FICHA_CONSUMO.TextBox1.Select()
            Me.Close()
        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 100
                DataGridView1.Columns(1).Width = 180
                DataGridView1.Columns(2).Width = 200
                DataGridView1.Columns(3).Width = 170
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class