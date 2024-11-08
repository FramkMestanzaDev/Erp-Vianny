Public Class Lista_Tela
    Dim dt As New DataTable
    Private Sub Lista_Tela_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim f As New fproductos
        Dim ol As String
        ol = Label2.Text
        dt = f.buscar_productos_almacen06(ol)

        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(2).Width = 540
            DataGridView1.Columns(3).Width = 130
            DataGridView1.Columns(4).Width = 100
            DataGridView1.Columns(5).Width = 100
            DataGridView1.Columns(6).Width = 200
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
            'DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(7).Visible = False
        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Producto" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 80
                DataGridView1.Columns(2).Width = 540
                DataGridView1.Columns(3).Width = 130
                DataGridView1.Columns(4).Width = 100
                DataGridView1.Columns(5).Width = 100
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = False
                'DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
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
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CODIGO" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 80
                DataGridView1.Columns(2).Width = 540
                DataGridView1.Columns(3).Width = 130
                DataGridView1.Columns(4).Width = 100
                DataGridView1.Columns(5).Width = 100
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = False
                'DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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

                    Orden_Produccion.TextBox4.Text = DataGridView1.Rows(a).Cells(6).Value
                    Orden_Produccion.TextBox12.Text = DataGridView1.Rows(a).Cells(2).Value
                    Orden_Produccion.TextBox16.Text = DataGridView1.Rows(a).Cells(3).Value
                    Orden_Produccion.TextBox17.Text = DataGridView1.Rows(a).Cells(4).Value
                End If
            Next
            Close()
        End If


    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar1()
    End Sub
End Class