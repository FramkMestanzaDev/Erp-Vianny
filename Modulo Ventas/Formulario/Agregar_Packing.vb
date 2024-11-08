Public Class Agregar_Packing
    Dim dt As New DataTable
    Private Sub Agregar_Packing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hj As New fingresopac
        Dim kk As New vpackingtela
        Dim j As Integer
        kk.gnumero_rollos = TextBox1.Text
        kk.galmac = Label3.Text
        dt = hj.buscar_packing(kk)

        If dt.Rows.Count > 0 Then
            DataGridView2.DataSource = dt
            j = DataGridView2.Rows.Count
            DataGridView1.Rows.Add(j)

            For i = 0 To j - 1
                DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(i).Cells(0).Value
                DataGridView1.Rows(i).Cells(2).Value = DataGridView2.Rows(i).Cells(6).Value
                DataGridView1.Rows(i).Cells(3).Value = DataGridView2.Rows(i).Cells(1).Value
                DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(9).Value
                DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(8).Value
                DataGridView1.Rows(i).Cells(6).Value = DataGridView2.Rows(i).Cells(4).Value
            Next

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, I2, SUMA As Integer
        Dim HJ1 As New VSTOCK_PAR
        Dim TG1 As New FSTOCK_PAR
        i = DataGridView1.Rows.Count
        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                NotaIngreso.DataGridView1.Rows.Add(1)
                I2 = NotaIngreso.DataGridView1.Rows.Count
                If I2 = 1 Then
                    HJ1.gpo = Mid(DataGridView1.Rows(a).Cells(1).Value, 1, 6)
                    HJ1.gCCIA = Trim(Label4.Text)
                    NotaIngreso.DataGridView1.Rows(0).Cells(0).Value = 1
                    NotaIngreso.DataGridView1.Rows(0).Cells(1).Value = TG1.mostrar_op(HJ1)
                    NotaIngreso.DataGridView1.Rows(0).Cells(7).Value = Mid(DataGridView1.Rows(a).Cells(1).Value, 1, 6)
                    NotaIngreso.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(a).Cells(2).Value
                    NotaIngreso.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(a).Cells(3).Value
                    NotaIngreso.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(a).Cells(4).Value
                    NotaIngreso.DataGridView1.Rows(0).Cells(5).Value = DataGridView1.Rows(a).Cells(5).Value
                    NotaIngreso.DataGridView1.Rows(0).Cells(6).Value = "RLL"
                    NotaIngreso.DataGridView1.Rows(0).Cells(8).Value = Mid(DataGridView1.Rows(a).Cells(1).Value, 7, 1)
                Else
                    If I2 > 1 Then
                        HJ1.gpo = Mid(DataGridView1.Rows(a).Cells(1).Value, 1, 6)
                        HJ1.gCCIA = Trim(Label4.Text)
                        SUMA = NotaIngreso.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                        NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(0).Value = SUMA
                        NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(1).Value = TG1.mostrar_op(HJ1)
                        NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(7).Value = Mid(DataGridView1.Rows(a).Cells(1).Value, 1, 6)
                        NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(a).Cells(2).Value
                        NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(a).Cells(3).Value
                        NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(a).Cells(4).Value
                        NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(5).Value = DataGridView1.Rows(a).Cells(5).Value
                        NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(6).Value = "RLL"
                        NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(8).Value = Mid(DataGridView1.Rows(a).Cells(1).Value, 7, 1)
                    End If
                End If

            End If
        Next
        NotaIngreso.Button2.Enabled = True
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar_PARTIDA()
    End Sub

    Private Sub buscar_PARTIDA()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))

            dv.RowFilter = "PARTIDA" & " like '%" & TextBox2.Text & "%'"
            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class