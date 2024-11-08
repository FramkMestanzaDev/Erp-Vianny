Public Class FICHA_ALMACE
    Private Sub FICHA_ALMACE_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add(2)
        DataGridView1.Rows(1).Cells(0).Value = "08"
        DataGridView1.Rows(1).Cells(1).Value = "TELA PLANA "
        DataGridView1.Rows(0).Cells(0).Value = "06"
        DataGridView1.Rows(0).Cells(1).Value = "TELA PUNTO"
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        FICHA_TECNICA.Close()
        FICHA_TECNICA.Label39.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        FICHA_TECNICA.Label41.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        FICHA_TECNICA.ShowDialog()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class