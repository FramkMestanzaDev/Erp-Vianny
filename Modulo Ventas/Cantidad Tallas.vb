Public Class Cantidad_Tallas
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        For i = 1 To 10
            Op_Manufactura.DataGridView2.Rows(1).Cells(i).Value = DataGridView1.Rows(i - 1).Cells(1).Value
            Op_Manufactura.DataGridView4.Rows(1).Cells(i).Value = DataGridView1.Rows(i - 1).Cells(2).Value
        Next

        Dim at, valor, valor1 As New Integer

        at = DataGridView1.Rows.Count
        valor = 0
        valor1 = 0
        For i = 0 To at - 1
            valor = valor + DataGridView1.Rows(i).Cells(1).Value
            valor1 = valor1 + DataGridView1.Rows(i).Cells(2).Value
        Next
        TextBox1.Text = valor
        Op_Manufactura.TextBox21.Text = valor
        Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim at, valor, valor1 As New Integer

        at = DataGridView1.Rows.Count
        valor = 0
        valor1 = 0
        For i = 0 To at - 1
            valor = valor + DataGridView1.Rows(i).Cells(1).Value
            valor1 = valor1 + DataGridView1.Rows(i).Cells(2).Value
        Next
        TextBox1.Text = valor
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim at, valor, valor1 As New Integer

        at = DataGridView1.Rows.Count
        valor = 0
        valor1 = 0
        For i = 0 To at - 1
            valor = valor + DataGridView1.Rows(i).Cells(1).Value
            valor1 = valor1 + DataGridView1.Rows(i).Cells(2).Value
        Next
        TextBox1.Text = valor
    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        Dim at, valor, valor1 As New Integer

        at = DataGridView1.Rows.Count
        valor = 0
        valor1 = 0
        For i = 0 To at - 1
            valor = valor + DataGridView1.Rows(i).Cells(1).Value
            valor1 = valor1 + DataGridView1.Rows(i).Cells(2).Value
        Next
        TextBox1.Text = valor
    End Sub

    Private Sub Cantidad_Tallas_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class