Imports System.Data.SqlClient
Imports System.Windows.Forms
Public Class TodOds
    Dim da As New DataTable
    Public conx As SqlConnection
    Dim Rsr212 As SqlDataReader
    Dim CodVen As String
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub MOSTRAR()
        DataGridView1.DataSource = Nothing
        abrir()
        Dim cmd As New SqlDataAdapter("select distinct  program_3 as PM, c.nomb_10 as CLIENTE from custom_vianny.dbo.qag0300 q inner join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 =c.fich_10
 where q.ccia ='01' and len(program_3) >0  and year( q.fcom_3) >='2023' and tipped_3 in ('M','D')order by nomb_10", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        If da.Rows.Count > 0 Then
            DataGridView1.Columns(0).Width = 100
            DataGridView1.Columns(1).Width = 400
        End If
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            Dim K As Integer


            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "CLIENTE" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 100
                DataGridView1.Columns(1).Width = 400
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar1()
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

            dv.RowFilter = "PM" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 100
                DataGridView1.Columns(1).Width = 400
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TodOds_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        DataGridView1.DataSource = ""
        da.Clear()
        MOSTRAR()
        TextBox1.Select()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar1()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar2()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        OdxPm.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Me.Close()
    End Sub
End Class