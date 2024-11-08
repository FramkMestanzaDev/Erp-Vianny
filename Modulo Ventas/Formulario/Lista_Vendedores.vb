Public Class Lista_Vendedores
    Private Sub Lista_Vendedores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add(3)
        DataGridView1.Rows(0).Cells(0).Value = "0001"
        DataGridView1.Rows(0).Cells(1).Value = "VSILVERIO"
        DataGridView1.Rows(1).Cells(0).Value = "0003"
        DataGridView1.Rows(1).Cells(1).Value = "GBALVIN"
        DataGridView1.Rows(2).Cells(0).Value = "0025"
        DataGridView1.Rows(2).Cells(1).Value = "VGRAUS"
        DataGridView1.Rows(2).Cells(0).Value = "0031"
        DataGridView1.Rows(2).Cells(1).Value = "JPOZZI"
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        OD.TextBox15.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        OD.TextBox78.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Close()
    End Sub
End Class