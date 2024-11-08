Public Class Form_Cotizacion3
    Public padre As New CotizacionUdp
    Private Sub Form_Cotizacion3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim jh As Integer
        'DataGridView1.Rows.Clear()
        'jh = HojaCotizacion.DataGridView6.Rows.Count
        'If jh = 0 Then

        'End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim va, la As Integer
        Dim sum As Double
        va = DataGridView1.Rows.Count
        padre.DataGridView6.Rows.Clear()
        If va > 0 Then
            padre.DataGridView6.Rows.Add(va)
            For i = 0 To va - 1
                padre.DataGridView6.Rows(i).Cells(0).Value = DataGridView1.Rows(i).Cells(0).Value
                padre.DataGridView6.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(1).Value
                padre.DataGridView6.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(2).Value
                padre.DataGridView6.Rows(i).Cells(3).Value = DataGridView1.Rows(i).Cells(3).Value
                padre.DataGridView6.Rows(i).Cells(4).Value = DataGridView1.Rows(i).Cells(4).Value
                padre.DataGridView6.Rows(i).Cells(5).Value = DataGridView1.Rows(i).Cells(5).Value


            Next
        End If
        la = DataGridView1.Rows.Count
        For p = 0 To la - 1
            sum = sum + CDbl(DataGridView1.Rows(p).Cells(5).Value)
        Next
        padre.TextBox21.Text = Math.Round(sum, 3)

        Close()

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, fila As Integer
        'Dim cant9 As Double
        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CONSUMO" Then
            DataGridView1.Rows(fila).Cells(3).Value = (Val(DataGridView1.Rows(fila).Cells(1).Value))
            Me.DataGridView1.Columns(3).DefaultCellStyle.Format = "n3"
            DataGridView1.Rows(fila).Cells(5).Value = Val(DataGridView1.Rows(fila).Cells(3).Value) * (Val(DataGridView1.Rows(fila).Cells(4).Value))
            Me.DataGridView1.Columns(5).DefaultCellStyle.Format = "n3"
        End If
        'If DataGridView1.Columns(e.ColumnIndex).HeaderText = "% MERMA" Then
        '    DataGridView1.Rows(fila).Cells(4).Value = (Val(DataGridView1.Rows(fila).Cells(2).Value))
        '    Me.DataGridView1.Columns(4).DefaultCellStyle.Format = "n3"
        '    DataGridView1.Rows(fila).Cells(6).Value = Val(DataGridView1.Rows(fila).Cells(4).Value) * (Val(DataGridView1.Rows(fila).Cells(5).Value))
        '    Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n3"
        'End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "C.UNITARIO" Then
            DataGridView1.Rows(fila).Cells(5).Value = Val(DataGridView1.Rows(fila).Cells(3).Value) * (Val(DataGridView1.Rows(fila).Cells(4).Value))
            Me.DataGridView1.Columns(5).DefaultCellStyle.Format = "n3"

        End If
    End Sub
End Class