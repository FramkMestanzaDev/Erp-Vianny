Public Class Detalle_Factura
    Dim da As New DataTable
    Private Sub Detalle_Factura_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fu As New freportedoc
        Dim fu1 As New vreportedoc

        fu1.gfactura2 = TextBox1.Text

        da = fu.reportedoc6(fu1)
        If da.Rows.Count <> 0 Then
            DataGridView1.DataSource = da

            DataGridView1.Columns(0).Width = 90
            DataGridView1.Columns(1).Width = 117
            DataGridView1.Columns(2).Width = 120
            DataGridView1.Columns(3).Width = 130
            DataGridView1.Columns(4).Width = 130
            DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns(1).DefaultCellStyle.BackColor = Color.Khaki
            DataGridView1.Columns(2).DefaultCellStyle.BackColor = Color.Beige
        End If
    End Sub
End Class