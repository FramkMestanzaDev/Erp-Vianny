Public Class Reporte_Facturas
    Dim dt As New DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.DataSource = ""
        Dim gh As New freportedoc
        Dim gh1 As New vreportedoc

        gh1.gfechaini = DateTimePicker1.Value
        dt = gh.reportedoc3(gh1)

        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(1).Width = 90
            DataGridView1.Columns(2).Width = 80
            DataGridView1.Columns(3).Width = 120
            DataGridView1.Columns(4).Width = 130
            DataGridView1.Columns(5).Width = 130
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns(3).DefaultCellStyle.BackColor = Color.Khaki
            'DataGridView1.Columns(5).DefaultCellStyle.ForeColor = Color.White
        End If

    End Sub

    Private Sub Reporte_Facturas_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Buscar_Factura.Show()
        End If
    End Sub
    Dim dy As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("EL CAMPO FACTURA NO DEBE ESTAR VACIO")
        Else
            DataGridView1.DataSource = ""
            Dim fu As New freportedoc
            Dim fu1 As New vreportedoc

            fu1.gfactura = TextBox1.Text
            dy = fu.reportedoc4(fu1)
            If dy.Rows.Count <> 0 Then
                DataGridView1.DataSource = dy
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(2).Width = 117
                DataGridView1.Columns(3).Width = 120
                DataGridView1.Columns(4).Width = 130
                DataGridView1.Columns(5).Width = 130
                DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                DataGridView1.Columns(2).DefaultCellStyle.BackColor = Color.Khaki
                DataGridView1.Columns(3).DefaultCellStyle.BackColor = Color.Beige
            End If
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString

            Detalle_Factura.TextBox1.Text = DataGridView1.Item(5, num1).Value()
            Detalle_Factura.Show()
        End If
    End Sub
End Class