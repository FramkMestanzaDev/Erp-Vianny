Public Class Partidas
    Dim dt4 As DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim jh As New vpartida
        Dim jh1 As New fpartida
        Dim gy As New vpackingtela
        Dim gy1 As New fingresopac
        Dim i, i2, i3, a, conta As Integer
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

            For a = 0 To i - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then

                    jh.gpartida = DataGridView1.Rows(a).Cells(1).Value
                    jh.galmacen = Packing.TextBox12.Text
                    gy.gcodigo_det = DataGridView1.Rows(a).Cells(5).Value
                    gy.ges = 1
                    gy1.actualizar_seleccionado(gy)
                    dt4 = jh1.buscar_partida2(jh)

                    If dt4.Rows.Count <> 0 Then
                        Packing.DataGridView2.DataSource = dt4
                        i2 = Packing.DataGridView2.Rows.Count
                        i3 = Packing.DataGridView1.Rows.Count

                        For a1 = 0 To i2 - 1
                            Packing.DataGridView1.Rows.Add()
                            Packing.DataGridView1.Rows(a1 + i3).Cells(1).Value = Packing.DataGridView2.Rows(a1).Cells(0).Value
                            Packing.DataGridView1.Rows(a1 + i3).Cells(2).Value = Packing.DataGridView2.Rows(a1).Cells(1).Value
                            Packing.DataGridView1.Rows(a1 + i3).Cells(3).Value = Packing.DataGridView2.Rows(a1).Cells(2).Value
                            Packing.DataGridView1.Rows(a1 + i3).Cells(4).Value = Packing.DataGridView2.Rows(a1).Cells(3).Value
                            Packing.DataGridView1.Rows(a1 + i3).Cells(5).Value = jh.gpartida
                            Packing.DataGridView1.Rows(a1 + i3).Cells(0).Value = True
                            Packing.DataGridView1.Rows(a1 + i3).Cells(6).Value = Packing.DataGridView2.Rows(a1).Cells(4).Value
                            Packing.DataGridView1.Rows(a1 + i3).Cells(7).Value = DataGridView1.Rows(a).Cells(5).Value
                            gy.gcodigo_det = Packing.DataGridView2.Rows(a1).Cells(4).Value
                            gy.ges = 1
                            gy1.actualizar_seleccionado_detalle(gy)
                        Next


                    End If

                End If
            Next

            Close()
        End If





    End Sub

    Private Sub Partidas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim k As Integer
        k = DataGridView1.Rows.Count

        For i = 0 To k - 1
            If DataGridView1.Rows(i).Cells(6).Value = 1 Then

                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.AliceBlue
                DataGridView1.Rows(i).ReadOnly = True
            End If
        Next
        'DataGridView1.Columns(5).Visible = False
        'DataGridView1.Columns(6).Visible = False
    End Sub
End Class