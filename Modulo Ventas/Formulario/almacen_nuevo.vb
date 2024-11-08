Imports System.Data.SqlClient
Public Class almacen_nuevo
    Sub informacionFactura()
        Dim cmd As New SqlDataAdapter("select ccia,talm_15m AS ALMACEN,nomb_15m AS DESCRIPCION  from custom_vianny.dbo.Mag1500 where ccia ='" + Label2.Text + "' and seriencr_15m ='1' and tienda_15m ='1'", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(2).Width = 350
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.BackColor = Color.AliceBlue
            DataGridView1.Columns(2).DefaultCellStyle.ForeColor = Color.Red
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub
    Sub infromacionRegistroProductos()
        Dim cmd As New SqlDataAdapter("select ccia,talm_15m AS ALMACEN,nomb_15m AS DESCRIPCION  from custom_vianny.dbo.Mag1500 where ccia ='" + Label2.Text + "' and tienda_15m ='1'
", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(2).Width = 350
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.BackColor = Color.AliceBlue
            DataGridView1.Columns(2).DefaultCellStyle.ForeColor = Color.Red
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub
    Private Sub almacen_nuevo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        If Label4.Text = "1" Then
            infromacionRegistroProductos()
        Else
            informacionFactura()
        End If
    End Sub
    Public conx As SqlConnection
    Public conn As SqlDataAdapter

    Public comando As SqlCommand
    Dim da As New DataTable
    Dim bs As New BindingSource()
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub



    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Label4.Text = "0" Then
            Registro_Facturas.Close()
            Registro_Facturas.Label27.Text = Label3.Text
            Registro_Facturas.TextBox8.Text = "ALMACEN" & " " & Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) & "-  REGISTRO DE VENTA  "
            Registro_Facturas.Label5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Registro_Facturas.Label29.Text = Label2.Text
            Registro_Facturas.TextBox10.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Registro_Facturas.Show()
        Else
            Dim CP As New Codificacion_Prodcutos
            CP.Label13.Text = Label2.Text
            CP.Label14.Text = "2"
            CP.Label20.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString().Trim()
            CP.Label19.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
            CP.ShowDialog()
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class