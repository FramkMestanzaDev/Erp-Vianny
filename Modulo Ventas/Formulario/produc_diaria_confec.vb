Public Class produc_diaria_confec
    Dim dt, dT1, dt10 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fg As New forden_costura
        Dim fg1 As New vorden_costura
        Dim U, U2, U3, kl As Integer
        DataGridView1.DataSource = ""
        DataGridView4.DataSource = ""
        fg1.gop = TextBox1.Text
        dt = fg.fecha_produc_diaria(fg1)
        dt10 = fg.fecha_produc_diaria_10(fg1)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView4.DataSource = dt10
            U = DataGridView1.Rows.Count
            kl = DataGridView4.Rows.Count
            DataGridView2.Columns.Add(DataGridView1.Rows(0).Cells(0).Value, DataGridView1.Rows(0).Cells(0).Value)
            For t = 0 To kl - 1
                DataGridView2.Rows.Add(DataGridView4.Rows(t).Cells(0).Value)
            Next


            For i = 0 To U - 1

                If i < U - 1 Then
                    If DataGridView1.Rows(i).Cells(0).Value = DataGridView1.Rows(i + 1).Cells(0).Value Then
                    Else
                        DataGridView2.Columns.Add(DataGridView1.Rows(i + 1).Cells(0).Value, DataGridView1.Rows(i + 1).Cells(0).Value)

                    End If
                End If



            Next


            U2 = DataGridView2.Rows.Count
            U3 = Integer.Parse(DataGridView2.ColumnCount.ToString())
            For i4 = 0 To U - 1
                For i3 = 0 To U2 - 1
                    For i5 = 1 To U3 - 1
                        'MsgBox(DataGridView1.Rows(i4).Cells(0).Value)
                        'MsgBox(DataGridView2.Columns(i5).HeaderText)
                        If DataGridView1.Rows(i4).Cells(1).Value.ToString.Trim = DataGridView2.Rows(i3).Cells(0).Value.ToString.Trim And DataGridView1.Rows(i4).Cells(0).Value = DataGridView2.Columns(i5).HeaderText Then

                            DataGridView2.Rows(i3).Cells(i5).Value = DataGridView1.Rows(i4).Cells(2).Value
                        End If
                    Next


                Next
            Next
            DataGridView2.Columns(0).Frozen = True

        Else
            MsgBox("NO HAY REGISTRO DE CONFECCION EN LA FECHA INDICADA")
            DataGridView1.DataSource = ""
        End If

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim fg As New forden_costura
        Dim fg1 As New vorden_costura

        DataGridView3.DataSource = ""
        fg1.gfecha = DateTimePicker1.Value
        dT1 = fg.fecha_produc_diaria1(fg1)
        If dT1.Rows.Count <> 0 Then
            DataGridView3.DataSource = dT1

            DataGridView3.Columns(1).Width = 350
            DataGridView3.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView3.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView3.Columns(0).Visible = False
        Else
            MsgBox("NO EXISTE INFORMACION EN LA FECHA SOLICTADA")
            DataGridView1.DataSource = ""
        End If


    End Sub
End Class