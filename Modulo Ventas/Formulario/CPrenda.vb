Public Class CPrenda
    Dim DT As DataTable
    Private Sub CPrenda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim f As New combiancines
        Dim f1 As New vproducot
        TextBox1.Text = ""
        Label1.Text = "COLOR"
        TextBox1.Select()
        f1.gcia = Label2.Text
        DT = f.mostrar_colores(f1)

        If DT.Rows.Count <> 0 Then
            DataGridView1.DataSource = DT
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(1).Width = 100
            DataGridView1.Columns(2).Width = 250
            DataGridView1.Columns(3).Width = 150
        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "DESCRIPCION" & " like '%" & Trim(TextBox1.Text) & "%'"

            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                If Label4.Text = 1 Then
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(3).Visible = False
                    DataGridView1.Columns(4).Visible = False
                    DataGridView1.Columns(5).Visible = False
                    DataGridView1.Columns(6).Visible = False
                    DataGridView1.Columns(1).Width = 100
                    DataGridView1.Columns(2).Width = 250
                    DataGridView1.Columns(3).Width = 150
                Else
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(1).Visible = False
                    DataGridView1.Columns(2).Visible = False
                    DataGridView1.Columns(3).Visible = False

                    DataGridView1.Columns(4).Width = 120
                    DataGridView1.Columns(5).Width = 200
                End If




            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'OD.TextBox19.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        'OD.TextBox59.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        'Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            With DataGridView1

                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                End If
                Op_Manufactura.TextBox17.Text = DataGridView1.Rows(hti.RowIndex).Cells(3).Value
                Op_Manufactura.colorprenda.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value



            End With
            Me.Close()
        End If
    End Sub
End Class