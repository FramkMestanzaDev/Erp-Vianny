Public Class Tabla_Articulos
    Dim DT As New DataTable
    Private Sub Tabla_Articulos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim HG As New fnia
        Dim hg1 As New vnia

        hg1.gccia = Label3.Text
        DT = HG.TABLA_PRODUCTOS(hg1)
        DataGridView1.DataSource = DT
        DataGridView1.Columns(1).Width = 180
        DataGridView1.Columns(2).Width = 450
    End Sub


    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CODIGO" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 180
                DataGridView1.Columns(2).Width = 450
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PRODUCTO" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 180
                DataGridView1.Columns(2).Width = 450
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar1()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a, conta As Integer
        i = DataGridView1.Rows.Count

        conta = 0



        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                conta = conta + 1
            End If
        Next
        If conta > 1 Then
            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
        Else

            For b = 0 To i - 1
                If Me.DataGridView1.Rows(b).Cells(0).Value = True Then

                    Reporte_Ingreso_Salida.TextBox2.Text = DataGridView1.Rows(b).Cells(1).Value
                    Reporte_Ingreso_Salida.TextBox3.Text = DataGridView1.Rows(b).Cells(2).Value
                    Reporte_Ingreso_Salida.Label6.Text = DataGridView1.Rows(b).Cells(1).Value
                End If
            Next

            Close()
        End If
    End Sub
End Class