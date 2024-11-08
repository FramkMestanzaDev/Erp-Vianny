Public Class BDocumentoContable
    Dim dt As New DataTable
    Private Sub BDocumentoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim jk As New vwilder
        If TextBox3.Text = "2" Then
            dt = jk.buscar_contabilidad
            If dt.Rows.Count <> 0 Then
                DataGridView1.DataSource = dt

                DataGridView1.Columns(1).Visible = True
                DataGridView1.Columns(2).Visible = True
                DataGridView1.Columns(3).Visible = True
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = True
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                DataGridView1.Columns(15).Visible = False
                DataGridView1.Columns(16).Visible = False
                DataGridView1.Columns(17).Visible = False
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(2).Width = 55
                DataGridView1.Columns(3).Width = 90
                DataGridView1.Columns(5).Width = 130
                DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(5).HeaderText = "ESTADO"
                DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.Red
                DataGridView1.Columns(5).DefaultCellStyle.ForeColor = Color.White
                Dim r As New Integer
                r = DataGridView1.Rows.Count
                For i = 0 To r - 1
                    If DataGridView1.Rows(i).Cells(5).Value.ToString.Trim = 0 Then
                        DataGridView1.Rows(i).Cells(5).Value = "PENDIENTE"
                    End If

                Next
            End If
        Else

            dt = jk.buscar()
            If dt.Rows.Count <> 0 Then
                DataGridView1.DataSource = dt

                DataGridView1.Columns(1).Visible = True
                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = True
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(1).Width = 120
                DataGridView1.Columns(5).Width = 120
                DataGridView1.Columns(12).Width = 130
                DataGridView1.Columns(12).HeaderText = "ESTADO"

                DataGridView1.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(12).DefaultCellStyle.BackColor = Color.Red
                DataGridView1.Columns(12).DefaultCellStyle.ForeColor = Color.White
                Dim r As New Integer
                r = DataGridView1.Rows.Count
                For i = 0 To r - 1
                    If DataGridView1.Rows(i).Cells(12).Value.ToString.Trim = 0 Then
                        DataGridView1.Rows(i).Cells(12).Value = "PENDIENTE"
                    End If

                Next

            End If
        End If
    End Sub
    Private Sub buscar_familia()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "factura" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Visible = True
                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = True
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
                'DataGridView1.Columns(5).Width = 120
                'DataGridView1.Columns(6).Width = 200
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox3.Text = "2" Then

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

                        Tesoreria.TextBox5.Text = DataGridView1.Rows(a).Cells(13).Value
                        Tesoreria.TextBox6.Text = DataGridView1.Rows(a).Cells(15).Value
                        Tesoreria.TextBox7.Text = DataGridView1.Rows(a).Cells(11).Value
                        Tesoreria.TextBox8.Text = DataGridView1.Rows(a).Cells(9).Value
                        Tesoreria.TextBox1.Text = DataGridView1.Rows(a).Cells(1).Value
                        Tesoreria.TextBox2.Text = DataGridView1.Rows(a).Cells(2).Value
                        Tesoreria.TextBox3.Text = DataGridView1.Rows(a).Cells(3).Value
                    End If
                Next
                Tesoreria.TextBox3.ReadOnly = False
                Tesoreria.TextBox3.ReadOnly = False
                Tesoreria.Button1.Enabled = True
                Tesoreria.Button2.Enabled = True
                Tesoreria.Button3.Enabled = True
                Tesoreria.Button4.Enabled = True
                Close()
            End If
        Else

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

                        Contabilidad.TextBox5.Text = DataGridView1.Rows(a).Cells(8).Value
                        Contabilidad.TextBox6.Text = DataGridView1.Rows(a).Cells(10).Value
                        Contabilidad.TextBox7.Text = DataGridView1.Rows(a).Cells(6).Value
                        Contabilidad.TextBox8.Text = DataGridView1.Rows(a).Cells(4).Value
                        Contabilidad.TextBox1.Text = DataGridView1.Rows(a).Cells(5).Value
                    End If
                Next
                Contabilidad.TextBox3.ReadOnly = False
                Contabilidad.TextBox3.ReadOnly = False
                Contabilidad.Button1.Enabled = True
                Contabilidad.Button2.Enabled = True
                Contabilidad.Button3.Enabled = True
                Contabilidad.Button4.Enabled = True
                Close()
            End If
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar_familia()
    End Sub
End Class