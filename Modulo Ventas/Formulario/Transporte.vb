Public Class Transporte
    Dim dt2 As DataTable
    Private Sub Transporte_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim gf As New fnopedido
        Dim OK As New vnapedido
        OK.galmacen = "10"
        dt2 = gf.validar_almacen2(OK)
        If dt2.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt2
            Dim k As Integer
            k = DataGridView1.Rows.Count
            For i = 0 To k - 1
                DataGridView1.Columns.Item(13).Visible = False
                DataGridView1.Columns.Item(14).Visible = False
                DataGridView1.Columns.Item(15).Visible = False
                DataGridView1.Columns.Item(16).Visible = False
                DataGridView1.Columns.Item(4).Frozen = True
                DataGridView1.Columns.Item(5).Frozen = True
                DataGridView1.Columns.Item(6).Frozen = True
                DataGridView1.Rows(i).Cells(4).Value = 1


                DataGridView1.Columns.Item(5).Width = 80
                DataGridView1.Columns.Item(6).Width = 90
                DataGridView1.Columns.Item(7).Width = 140
                DataGridView1.Columns.Item(8).Width = 220
                DataGridView1.Columns.Item(9).Width = 70
                DataGridView1.Columns.Item(10).Width = 70
                DataGridView1.Columns.Item(11).Width = 200
                DataGridView1.Columns.Item(12).Width = 75
                DataGridView1.Columns.Item(17).Width = 200
                If DataGridView1.Rows(i).Cells(13).Value = 2 Then
                    DataGridView1.Rows(i).Cells(0).Value = True
                End If
                If DataGridView1.Rows(i).Cells(14).Value = 2 Then
                    DataGridView1.Rows(i).Cells(1).Value = True
                End If
                If DataGridView1.Rows(i).Cells(16).Value = 2 Then
                    DataGridView1.Rows(i).Cells(2).Value = True
                    DataGridView1.Rows(i).Cells(2).ReadOnly = True
                End If
            Next
        Else
            MsgBox("NO HAY INFORMACION PARA MOSTRAR")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim u As Integer
        u = DataGridView1.Rows.Count

        For i = 0 To u - 1
            If DataGridView1.Rows(i).Cells(3).Value = True And DataGridView1.Rows(i).Cells(4).Value = 1 Then
                DataGridView1.Rows(i).Cells(4).Value = 2
                DataGridView2.Rows.Add(1)
                DataGridView2.Rows(i).Cells(4).Value = DataGridView1.Rows(i).Cells(12).Value
                DataGridView2.Rows(i).Cells(5).Value = DataGridView1.Rows(i).Cells(8).Value
                DataGridView2.Rows(i).Cells(6).Value = DataGridView1.Rows(i).Cells(6).Value
                DataGridView2.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(9).Value
                DataGridView2.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(10).Value
                DataGridView2.Rows(i).Cells(9).Value = DataGridView1.Rows(i).Cells(11).Value
                DataGridView2.Rows(i).Cells(10).Value = DataGridView1.Rows(i).Cells(17).Value
                DataGridView2.Rows(i).Cells(11).Value = DataGridView1.Rows(i).Cells(15).Value
                DataGridView1.Rows(i).ReadOnly = True


            End If
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class