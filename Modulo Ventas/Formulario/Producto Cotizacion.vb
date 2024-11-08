Public Class Producto_Cotizacion
    Public Property Padre As Form_Cotizacion
    Dim dt As New DataTable
    Private Sub Producto_Cotizacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Form_Cotizacion.TextBox3.Text = "1" Then
            mostrar()
        Else
            mostrar2()
        End If

    End Sub
    Private Sub mostrar()
        Try
            Dim func As New fprducto



            dt = func.mostrar_tela
            If dt.Rows.Count <> 0 Then
                DataGridView1.DataSource = dt

                DataGridView1.Columns(0).Width = 150
                DataGridView1.Columns(1).Width = 503
                DataGridView1.Columns(2).Width = 100
            End If
        Catch ex As Exception

            MsgBox("NO EXISTE INFORMACION PARA MOSTRAR")

        End Try
    End Sub
    Private Sub mostrar2()
        Try
            Dim func As New fprducto



            dt = func.mostrar_product
            If dt.Rows.Count <> 0 Then
                DataGridView1.DataSource = dt

                DataGridView1.Columns(0).Width = 150
                DataGridView1.Columns(1).Width = 503
                DataGridView1.Columns(2).Width = 100
            End If
        Catch ex As Exception

            MsgBox("NO EXISTE INFORMACION PARA MOSTRAR")

        End Try
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Descripcion" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(0).Width = 150
                DataGridView1.Columns(1).Width = 503
                DataGridView1.Columns(2).Width = 100
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

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Padre.TextBox1.Text = Trim(DataGridView1.SelectedCells.Item(0).Value)
        Padre.TextBox2.Text = Trim(DataGridView1.SelectedCells.Item(1).Value)
        Padre.TextBox4.Text = Trim(DataGridView1.SelectedCells.Item(3).Value)
        Padre.Button1.Select()
        Close()

    End Sub
End Class