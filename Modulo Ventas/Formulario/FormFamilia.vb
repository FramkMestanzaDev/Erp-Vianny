Public Class FormFamilia
    Dim dt As DataTable
    Private Sub FormFamilia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox1.Select()
        DataGridView1.DataSource = Nothing
        Dim f As New ffamilia
        dt = f.mostrar_familia
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Width = 120
            DataGridView1.Columns(5).Width = 200
        End If

    End Sub
    Private Sub buscar_familia()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "DESCRIPCION" & " like '%" & Trim(TextBox1.Text) & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(1).Visible = False
                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(3).Visible = False

                DataGridView1.Columns(4).Width = 120
                DataGridView1.Columns(5).Width = 200
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar_familia()
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If Label2.Text = 1 Then
            Op_Manufactura.TextBox12.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value
            Op_Manufactura.TextBox27.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
            Op_Manufactura.TextBox28.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
            'Else
            '    OD.TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value
            '    OD.TextBox16.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
            '    OD.TextBox79.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        End If
        TextBox1.Text = ""
        Me.Close()


    End Sub


End Class