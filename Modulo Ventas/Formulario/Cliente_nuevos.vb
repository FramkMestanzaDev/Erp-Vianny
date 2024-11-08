Imports System.Data.SqlClient
Imports System.Windows.Forms.DataVisualization.Charting
Public Class Cliente_nuevos
    Dim DT, DT2 As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim ty, ty2, ty3 As New DataTable
    Public conn As SqlDataAdapter

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim df As New Exportar
        df.llenarExcel(DataGridView1)
    End Sub

    Private Sub Cliente_nuevos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Dim dt10 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        dt10.Clear()
        DataGridView1.DataSource = Nothing
        Dim cmd5 As New SqlDataAdapter("select  r_social AS CLIENTE, (select ISNULL( SUM(vvta1_3),0) from custom_vianny.dbo.fag0300 where fich_3 =COD_CLI  and  (fcom_3 >='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' and fcom_3 <='" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "') and flag_3 =1 ) as VENTA, alias_ven as VENDEDOR,fcom AS F_INICIO from cliente a
left join custom_vianny.dbo.Vendedores v on a.Vendedor = v.codigo_ven
where (fcom >='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' and fcom<='" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "')  group by alias_ven,r_social,COD_CLI,fcom order by  Vendedor", conx)
        cmd5.Fill(dt10)
        If dt10.Rows.Count > 0 Then
            DataGridView1.DataSource = dt10
            DataGridView1.Columns(0).Width = 270
            For i = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(1).Value <= 0 Then
                    Dim pos As Integer = i
                    Me.DataGridView1.CurrentCell = Nothing
                    Me.DataGridView1.Rows(pos).Visible = False
                End If
            Next
        Else
            DataGridView1.DataSource = Nothing
        End If



    End Sub
End Class