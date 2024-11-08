Public Class Form_Padron
    Dim DT, dt1 As New DataTable
    Sub mostrar()
        Dim f As New Padron_tallas
        Dim f1 As New vtallas
        f1.gccia = Label2.Text
        DT = f.mostrar_padrom_tallas(f1)

        If DT.Rows.Count <> 0 Then
            DataGridView1.DataSource = DT

            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(2).Width = 70
            DataGridView1.Columns(14).Width = 300



        End If
    End Sub
    Private Sub Form_Padron_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mostrar()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a, conta As Integer
        i = DataGridView1.Rows.Count
        a = 0
        conta = 0

        If Label3.Text = 1 Then
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

                        Op_Manufactura.TextBox14.Text = DataGridView1.Rows(a).Cells(2).Value
                        Op_Manufactura.TextBox15.Text = DataGridView1.Rows(a).Cells(14).Value.ToString.Trim
                        'DataGridView1.Rows(a).Cells(13).Value.ToString.Trim & "/" & DataGridView1.Rows(a).Cells(14).Value.ToString.Trim & "/" & DataGridView1.Rows(a).Cells(15).Value.ToString.Trim & "/" &
                        '     DataGridView1.Rows(a).Cells(16).Value.ToString.Trim & "/" & DataGridView1.Rows(a).Cells(17).Value.ToString.Trim & "/" & DataGridView1.Rows(a).Cells(18).Value.ToString.Trim & "/" & DataGridView1.Rows(a).Cells(19).Value.ToString.Trim & "/" & DataGridView1.Rows(a).Cells(20).Value.ToString.Trim & "/" & DataGridView1.Rows(a).Cells(21).Value.ToString.Trim & "/" & DataGridView1.Rows(a).Cells(22).Value.ToString.Trim
                        Dim ab As New vtallas
                        Dim ab1 As New Padron_tallas

                        ab.gcodigo = DataGridView1.Rows(a).Cells(2).Value.ToString.Trim
                        ab.gccia = Label2.Text
                        dt1 = ab1.mostrar_padrom_tallas_SELECCIONADO(ab)
                        If dt1.Rows.Count <> 0 Then
                            Op_Manufactura.DataGridView2.DataSource = dt1
                            Op_Manufactura.DataGridView4.DataSource = dt1
                        End If
                    End If
                Next
                Op_Manufactura.Button2.Enabled = True
                Op_Manufactura.TextBox21.Text = ""
                Close()

            End If
        Else
            If Label3.Text = 2 Then
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

                            OD.TextBox13.Text = DataGridView1.Rows(a).Cells(2).Value
                            OD.TextBox20.Text = DataGridView1.Rows(a).Cells(14).Value.ToString.Trim
                            Dim ab As New vtallas
                            Dim ab1 As New Padron_tallas

                            ab.gcodigo = DataGridView1.Rows(a).Cells(2).Value.ToString.Trim
                            ab.gccia = Label2.Text
                            dt1 = ab1.mostrar_padrom_tallas_SELECCIONADO(ab)
                            If dt1.Rows.Count <> 0 Then
                                OD.DataGridView5.DataSource = dt1
                                OD.DataGridView4.DataSource = dt1
                            End If
                        End If
                    Next

                    Close()

                End If
            End If
        End If



    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "TALLAS" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(1).Visible = False
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(2).Width = 70
                DataGridView1.Columns(14).Width = 300
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        PadronTalla.limpiar()
        PadronTalla.ShowDialog()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class