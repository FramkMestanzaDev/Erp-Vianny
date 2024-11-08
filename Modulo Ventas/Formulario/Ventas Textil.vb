Public Class Ventas_Textil
    Dim dt As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim vf As New fventas
        Dim vf1 As New vventas
        Dim IA As Integer
        vf1.gfini = DateTimePicker1.Value
        vf1.gffin = DateTimePicker2.Value

        dt = vf.buscar_ventas(vf1)
        If dt.Rows.Count <> 0 Then
            DataGridView2.DataSource = dt

            IA = DataGridView2.Rows.Count
            DataGridView1.Rows.Add(IA)

            For a = 0 To IA - 1
                DataGridView1.Rows(a).Cells(0).Value = DataGridView2.Rows(a).Cells(0).Value & "-" & DataGridView2.Rows(a).Cells(1).Value
                DataGridView1.Rows(a).Cells(1).Value = DataGridView2.Rows(a).Cells(8).Value
                DataGridView1.Rows(a).Cells(2).Value = DataGridView2.Rows(a).Cells(7).Value
                DataGridView1.Rows(a).Cells(3).Value = DataGridView2.Rows(a).Cells(2).Value
                DataGridView1.Rows(a).Cells(4).Value = DataGridView2.Rows(a).Cells(9).Value
                DataGridView1.Rows(a).Cells(5).Value = DataGridView2.Rows(a).Cells(12).Value
                DataGridView1.Rows(a).Cells(6).Value = DataGridView2.Rows(a).Cells(13).Value
                DataGridView1.Rows(a).Cells(7).Value = DataGridView2.Rows(a).Cells(14).Value
                DataGridView1.Rows(a).Cells(8).Value = DataGridView2.Rows(a).Cells(15).Value
                DataGridView1.Rows(a).Cells(9).Value = DataGridView2.Rows(a).Cells(17).Value
                If DataGridView2.Rows(a).Cells(17).Value = 69 Then
                    DataGridView1.Rows(a).Visible = False
                End If
            Next
        End If
    End Sub



    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, fila As Integer

        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PRECIO" Then



            DataGridView1.Rows(fila).Cells(8).Value = DataGridView1.Rows(fila).Cells(4).Value * DataGridView1.Rows(fila).Cells(5).Value
            DataGridView1.Rows(fila).Cells(6).Value = DataGridView1.Rows(fila).Cells(8).Value / 1.18
            DataGridView1.Rows(fila).Cells(7).Value = DataGridView1.Rows(fila).Cells(8).Value - DataGridView1.Rows(fila).Cells(7).Value

            Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
            Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
            Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n2"



        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Ventas_Textil_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class