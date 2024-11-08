Imports System.Data.SqlClient
Public Class Buzin_reclamos
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim da As New DataTable
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Buzin_reclamos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        abrir()
        Dim cmd As New SqlDataAdapter("select numero AS NUMERO,ruc AS RUC,cliente AS CLIENTE,nota_pedido AS NOTA_PEDIDO,ap_am AS CONTACTO,lugar_visita AS LUGAR_VISITA, telefono AS TELEFONO, fecha AS FECHA,vendedor from Hoja_Reclamo_tela", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(2).Width = 95
        DataGridView1.Columns(3).Width = 95
        DataGridView1.Columns(4).Width = 200
        DataGridView1.Columns(5).Width = 95
        DataGridView1.Columns(6).Width = 200
        DataGridView1.Columns(7).Width = 200
        DataGridView1.Columns(8).Width = 80
        DataGridView1.Columns(8).Visible = False
        DataGridView1.Columns(9).Width = 80
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 1 Then
            Rpt_Reclamo.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
            Rpt_Reclamo.Show()
        Else
            If e.ColumnIndex = 0 Then
                Reclamos_calidad.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
                Reclamos_calidad.Label2.Text = DataGridView1.Rows(e.RowIndex).Cells(8).Value
                Reclamos_calidad.Show()
            End If
        End If

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub
End Class