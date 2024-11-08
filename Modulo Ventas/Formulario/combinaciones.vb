Public Class combinaciones
    Dim DT As DataTable
    Dim DT1 As DataTable
    Private Sub Combinaciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load




        TextBox1.Text = ""
                Label1.Text = "FAMILIA"
                TextBox1.Select()
                Dim f As New ffamilia
                DT = f.mostrar_familia
        If DT.Rows.Count <> 0 Then
            DataGridView1.DataSource = DT
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Width = 120
            DataGridView1.Columns(5).Width = 200
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub




    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub



    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            With DataGridView1

                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                End If

                If Label4.Text = 1 Then
                    If Label3.Text = 1 Then
                        Op_Manufactura.TextBox18.Text = DataGridView1.Rows(hti.RowIndex).Cells(4).Value
                        Op_Manufactura.colortela.Text = DataGridView1.Rows(hti.RowIndex).Cells(3).Value
                        Op_Manufactura.TextBox29.Text = DataGridView1.Rows(hti.RowIndex).Cells(2).Value

                    Else
                        OD.TextBox19.Text = DataGridView1.Rows(hti.RowIndex).Cells(3).Value
                        OD.TextBox59.Text = DataGridView1.Rows(hti.RowIndex).Cells(2).Value
                    End If
                Else

                    If Trim(Label3.Text) = 1 Then
                        Op_Manufactura.TextBox12.Text = DataGridView1.Rows(hti.RowIndex).Cells(4).Value
                        Op_Manufactura.TextBox27.Text = DataGridView1.Rows(hti.RowIndex).Cells(3).Value
                        Op_Manufactura.TextBox28.Text = DataGridView1.Rows(hti.RowIndex).Cells(2).Value
                    Else
                        If Trim(Label3.Text) = "2" Then

                            OD.TextBox11.Text = DataGridView1.Rows(hti.RowIndex).Cells(4).Value
                            OD.TextBox16.Text = DataGridView1.Rows(hti.RowIndex).Cells(3).Value
                            OD.TextBox79.Text = DataGridView1.Rows(hti.RowIndex).Cells(2).Value
                        End If

                    End If

                End If



            End With
            Me.Close()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class