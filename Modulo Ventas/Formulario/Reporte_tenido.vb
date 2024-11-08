Public Class Reporte_tenido
    Dim dt As New DataTable
    Private Sub Reporte_tenido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hj As New ftenido
        Dim AD1 As Integer



        dt = hj.GRAFICO_TENIDO()
        DataGridView1.DataSource = dt

        Me.Chart1.Titles.Add("DIAGRAMA DE DISPERCION")

        AD1 = DataGridView1.Rows.Count
        MsgBox(AD1)
        Chart1.ChartAreas.Add("VENTAS")
        Chart1.Series(0)("DrawingStyle") = "Cylinder"
        For i = 0 To AD1 - 1
            Me.Chart1.Series(0).Points.AddXY(DataGridView1.Rows(i).Cells(0).Value, DataGridView1.Rows(i).Cells(1).Value)
            Me.Chart1.Series(1).Points.AddXY(DataGridView1.Rows(i).Cells(0).Value, DataGridView1.Rows(i).Cells(2).Value)
            Me.Chart1.Series(2).Points.AddXY(DataGridView1.Rows(i).Cells(0).Value, DataGridView1.Rows(i).Cells(3).Value)
        Next
    End Sub
End Class