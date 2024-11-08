Imports System.Data.SqlClient
Public Class Factura_Letra
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
    Private Sub Factura_Letra_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT sfactu_3 as Serie,nfactu_3 as Correlativo, nomb_3 as Cliente, CASE when mone_3 =1 then 'S/' else '$' end as Moneda, case when mone_3=1 then tot1_3 else tot2_3 end as Total,fich_3  FROM custom_vianny.dbo.fag0300 WHERE ccia_3 ='01'  AND YEAR(fcom_3) >='2023' and tienda_3 <> '01'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 71
        DataGridView1.Columns(1).Width = 70
        DataGridView1.Columns(2).Width = 320
        DataGridView1.Columns(3).Width = 70
        DataGridView1.Columns(4).Width = 70
        DataGridView1.Columns(5).Visible = False
        TextBox1.Select()
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick



        Letra.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value) & "-" & Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        Letra.ComboBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
        Letra.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value)
        Letra.Label21.Text = 1
        Letra.TextBox8.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
        Letra.TextBox11.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
        TextBox1.Text = ""
        TextBox2.Text = ""
        Close()
    End Sub
    Private Sub buscarnombre()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Cliente" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 71
                DataGridView1.Columns(1).Width = 70
                DataGridView1.Columns(2).Width = 320
                DataGridView1.Columns(3).Width = 70
                DataGridView1.Columns(4).Width = 70

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscarserie()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Correlativo" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 71
                DataGridView1.Columns(1).Width = 70
                DataGridView1.Columns(2).Width = 320
                DataGridView1.Columns(3).Width = 70
                DataGridView1.Columns(4).Width = 70

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscarnombre()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscarserie()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class