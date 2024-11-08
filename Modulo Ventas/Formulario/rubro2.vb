Imports System.Data.SqlClient
Public Class rubro2
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
    Private Sub rubro2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT oper_14f AS RUBRO,nomb_14f as DESCRIPCION FROM custom_vianny.DBO.fag1400 WHERE  ccia_14f='" + Label1.Text + "' AND VEN=1", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).Width = 350
            DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).DefaultCellStyle.BackColor = Color.Black
            DataGridView1.Columns(1).DefaultCellStyle.ForeColor = Color.White
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub



    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Registro_Facturas.Close()
        Registro_Facturas.Label4.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Registro_Facturas.TextBox13.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Registro_Facturas.Label27.Text = Label3.Text
        Registro_Facturas.TextBox8.Text = "ALMACEN" & " " & Label4.Text & "-  REGISTRO DE VENTA  "
        Registro_Facturas.Label5.Text = Label4.Text
        Registro_Facturas.Label29.Text = Label1.Text
        Registro_Facturas.TextBox10.Text = Label4.Text
        Registro_Facturas.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class