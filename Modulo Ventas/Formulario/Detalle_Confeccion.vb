Public Class Detalle_Confeccion
    Dim dt As New DataTable
    Private Sub Detalle_Confeccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim gh As New forden_costura
        Dim jh As New vorden_costura
        Dim U As Integer
        Dim suma As Double
        jh.gorden = TextBox1.Text
        dt = gh.detalle_costura(jh)
        DataGridView1.DataSource = dt
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(2).Visible = False
        DataGridView1.Columns(4).Visible = False
        DataGridView1.Columns(6).Visible = False
        DataGridView1.Columns(3).HeaderText = "CORTE"
        DataGridView1.Columns(5).HeaderText = "PRENDAS TERMINADAS"
        DataGridView1.Columns(7).HeaderText = "FECHA/HORA REGISTRO"
        DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(3).Width = 100
        DataGridView1.Columns(5).Width = 150
        DataGridView1.Columns(7).Width = 150
        U = DataGridView1.Rows.Count
        For i = 0 To U - 1
            suma = suma + DataGridView1.Rows(i).Cells(5).Value
        Next
        TextBox4.Text = suma
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class