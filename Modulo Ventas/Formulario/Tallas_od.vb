Public Class Tallas_od
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For i = 1 To 10
            OD.DataGridView4.Rows(1).Cells(i).Value = DataGridView1.Rows(i - 1).Cells(1).Value
            OD.DataGridView5.Rows(1).Cells(i).Value = DataGridView1.Rows(i - 1).Cells(2).Value
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
        OD.TextBox69.Text = valor
        Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

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

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        ' Verificar si la celda actual es de tipo DataGridViewTextBoxCell
        If DataGridView1.CurrentCell.GetType() Is GetType(DataGridViewTextBoxCell) Then
            ' Obtener el control de edición asociado a la celda actual
            Dim textBox As TextBox = CType(e.Control, TextBox)

            ' Asociar el evento KeyPress para validar la entrada
            AddHandler textBox.KeyPress, AddressOf TextBox_KeyPress
        End If
    End Sub

    Private Sub TextBox_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        ' Verificar si la tecla presionada es un número o la tecla de retroceso
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Si no es un número y no es la tecla de retroceso, entonces no permitir la entrada
            e.Handled = True
        End If
    End Sub

    Private Sub Tallas_od_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ca As Integer

        ca = DataGridView1.Rows.Count

        For i = 0 To ca - 1

            If Convert.ToInt32(Trim(DataGridView1.Rows(i).Cells(0).Value).Length) = 0 Then

                DataGridView1.Rows(i).ReadOnly = True
            Else
                DataGridView1.Rows(i).ReadOnly = False
            End If
        Next
    End Sub
End Class