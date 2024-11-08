Public Class Aprobar_Cotizacion
    Dim DT, DT2, DT3, DT4, DT5 As New DataTable

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i As Integer
        Dim func As New VCOTIZACION
        Dim fg As New FCOTIZACION
        i = DataGridView1.Rows.Count
        If TextBox5.Text = 2 Then
            DataGridView1.Columns(0).Visible = False
            Button2.Visible = False
            Button3.Visible = False
        End If

        For i1 = 0 To i - 1
            If DataGridView1.Rows(i1).Cells(0).Value = True Then
                func.gid_cotizacion = DataGridView1.Rows(i1).Cells(7).Value
                func.gestado = 0
                fg.actualizar_ESTADO(func)
            End If
        Next
        MsgBox("SE REALIZO LA ACCION SOLICITADA")
        DT5 = fg.mostrar_COTIZACION4()
        DataGridView1.DataSource = DT5

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim i As Integer
        Dim func As New VCOTIZACION
        Dim fg As New FCOTIZACION
        i = DataGridView1.Rows.Count
        If TextBox5.Text = 2 Then
            DataGridView1.Columns(0).Visible = False
            Button2.Visible = False
            Button3.Visible = False
        End If
        DataGridView1.Columns(8).Visible = False
        For i1 = 0 To i - 1
            If DataGridView1.Rows(i1).Cells(0).Value = True Then
                func.gid_cotizacion = DataGridView1.Rows(i1).Cells(7).Value
                func.gestado = 1
                fg.actualizar_ESTADO(func)
            End If

        Next
        MsgBox("SE APROBO LA COTIZACION SOLICITADA")
        DT2 = fg.mostrar_COTIZACION3()

        DataGridView1.DataSource = DT2

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        buscar2()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim func10 As New FCOTIZACION

        If RadioButton2.Checked = True Then
            DT3 = func10.mostrar_COTIZACION4()
            If DT3.Rows.Count <> 0 Then
                DataGridView1.DataSource = DT3
                DataGridView1.Columns(8).Visible = False
                Dim i As Integer
                i = DataGridView1.Rows.Count

                For I1 = 0 To i - 1
                    If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                        DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.Red
                    End If

                Next
            Else
                DataGridView1.DataSource = ""
            End If
            Button3.Visible = True
            Button2.Visible = False
            If TextBox5.Text = 2 Then
                DataGridView1.Columns(0).Visible = False
                Button2.Visible = False
                Button3.Visible = False
            End If
        Else
            If RadioButton1.Checked = True Then
                DT4 = func10.mostrar_COTIZACION3()
                If DT4.Rows.Count <> 0 Then
                    DataGridView1.DataSource = DT4

                Else
                    DataGridView1.DataSource = ""
                End If
                Button2.Visible = True
                Button3.Visible = False
                If TextBox5.Text = 2 Then
                    DataGridView1.Columns(0).Visible = False
                    Button2.Visible = False
                    Button3.Visible = False
                End If
            End If
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Aprobar_Cotizacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim func As New FCOTIZACION
        RadioButton1.Checked = True

        DT = func.mostrar_COTIZACION3()
        If DT.Rows.Count <> 0 Then
            DataGridView1.DataSource = DT
            DataGridView1.Columns(8).Visible = False
            If TextBox5.Text = 2 Then
                DataGridView1.Columns(0).Visible = False
            End If
        End If
        If TextBox5.Text = 2 Then
            DataGridView1.Columns(0).Visible = False
            Button2.Visible = False
            Button3.Visible = False
        Else
            Button3.Visible = False
        End If

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "COTIZACION" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv


            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox4.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv


            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub
End Class