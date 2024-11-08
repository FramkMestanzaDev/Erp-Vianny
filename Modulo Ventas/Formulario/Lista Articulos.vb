Public Class Lista_Articulos
    Dim DT As DataTable
    Private Sub Lista_Articulos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim GH As New fcliente
        Dim jk As New vcliente


        jk.gccia = Trim(Label2.Text)


        DT = GH.buscar_codigo_almacen(jk)
        DataGridView1.DataSource = DT
        DataGridView1.Columns(1).Width = 135
        DataGridView1.Columns(2).Width = 350

    End Sub


    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 135
                DataGridView1.Columns(2).Width = 350

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

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a, conta As Integer
        i = DataGridView1.Rows.Count
        a = 0
        conta = 0



        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                conta = conta + 1
            End If
        Next
        If conta > 1 Then
            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
        Else


            For a = 0 To i - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then

                    NotaIngreso.TextBox20.Text = Trim(DataGridView1.Rows(a).Cells(1).Value)
                    NotaIngreso.TextBox20.Select()
                End If
            Next
            Close()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class