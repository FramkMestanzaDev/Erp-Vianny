Imports System.Data.SqlClient
Public Class TelaUdp
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Public comando As SqlCommand
    Dim da, da1 As New DataTable
    Dim Rsr2 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TelaUdp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True
        Listar_Informacion()
    End Sub
    Sub Listar_Informacion()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select FEnRtp as FEntergaCorte,OpRtp as Op,PoRtp as Po,ProRtp as Producto, CanRtp as Cantidad,case when EstRtp = '0' then 'PENDIENTE'  when EstRtp = '1' then 'DESPACHADO parcial' ELSE 'DESPACHADO' END as Estado,TallRtp as Taller,IdRtp,isnull(SelRtp,0) from RequerimientoTelaProd WHERE EstRtp  IN (0,1)", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(3).Width = 400
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(8).Visible = False

    End Sub
    Sub Listar_Informacion_Despachado()
        da1.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select FEnRtp as FEntergaCorte,OpRtp as Op,PoRtp as Po,ProRtp as Producto, CanRtp as Cantidad,case when EstRtp = '0' then 'PENDIENTE'  when EstRtp = '1' then 'DESPACHADO parcial' ELSE 'DESPACHADO' END as Estado,TallRtp as Taller,IdRtp,isnull(SelRtp,0) from RequerimientoTelaProd WHERE EstRtp  IN (2)", conx)
        cmd.Fill(da1)
        DataGridView1.DataSource = da1
        DataGridView1.Columns(3).Width = 400
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(8).Visible = False
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer
            dv.RowFilter = "OD" & " like '%" & TextBox1.Text & "%'"
            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(3).Width = 400
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim exportar As New Exportar
        exportar.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then
            da.Clear()
            da1.Clear()
            Listar_Informacion()
        Else
            da.Clear()
            da1.Clear()
            Listar_Informacion_Despachado()
        End If
    End Sub
End Class