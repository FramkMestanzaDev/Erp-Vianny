Public Class Procesos
    Dim ad As DataTable
    Private Sub Procesos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TG As New fproceso
        Dim TG1 As New vproceso
        Try

            TG1.gccia = Label2.Text
            ad = TG.bucar_proceso(TG1)
            DataGridView1.DataSource = ad
            DataGridView1.Columns(0).Width = 142
            DataGridView1.Columns(1).Width = 330
        Catch ex As Exception
            MsgBox("NO EXISTEN PROCESOS")
        End Try

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        NiaHilo.TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        NiaHilo.TextBox13.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If Label1.Text = 1 Then
            NiaHilo.TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            NiaHilo.TextBox13.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
            Me.Close()
        Else
            If Label1.Text = 2 Then
                NotaIngreso.TextBox14.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                NotaIngreso.TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                Me.Close()
            Else
                If Label1.Text = 3 Then
                    Nsalida.TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                    Nsalida.TextBox13.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                    Me.Close()
                Else
                    If Label1.Text = 4 Then
                        NsaHilo.TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                        NsaHilo.TextBox13.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                        Me.Close()
                    Else
                        If Label1.Text = 5 Then
                            os.TextBox14.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                            os.TextBox15.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                            Me.Close()
                        End If
                    End If
                End If
            End If
        End If

    End Sub
End Class