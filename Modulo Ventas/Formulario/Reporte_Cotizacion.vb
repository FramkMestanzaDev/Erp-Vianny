Public Class Reporte_Cotizacion
    Dim DT As New DataTable
    Private Sub Reporte_Cotizacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim FG As New FCOTIZACION

        DT = FG.mostrar_TODASCOTIZACIONES()

        If DT.Rows.Count <> 0 Then
            If TextBox4.Text = 1 Then
                DataGridView1.DataSource = DT
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(11).DefaultCellStyle.BackColor = Color.Bisque
                DataGridView1.Columns(12).DefaultCellStyle.BackColor = Color.LemonChiffon
                DataGridView1.Columns(14).DefaultCellStyle.BackColor = Color.DarkTurquoise
                DataGridView1.Columns(6).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(7).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(8).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(9).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(10).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(11).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(12).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(13).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(14).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(10).Width = 60
                DataGridView1.Columns(11).Width = 60
                DataGridView1.Columns(12).Width = 60
                DataGridView1.Columns(13).Width = 60
                DataGridView1.Columns(14).Width = 60
            Else
                DataGridView1.DataSource = DT
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(7).Width = 250
                DataGridView1.Columns(7).Width = 120
                DataGridView1.Columns(9).Width = 250

                DataGridView1.Columns(6).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(7).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(8).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(9).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                DataGridView1.Columns(15).Visible = False
            End If

        End If

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "COTIZACION" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(11).DefaultCellStyle.BackColor = Color.Bisque
                DataGridView1.Columns(12).DefaultCellStyle.BackColor = Color.LemonChiffon
                DataGridView1.Columns(14).DefaultCellStyle.BackColor = Color.DarkTurquoise
                DataGridView1.Columns(6).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(7).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(8).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(9).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(10).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(11).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(12).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(13).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(14).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(10).Width = 60
                DataGridView1.Columns(11).Width = 60
                DataGridView1.Columns(12).Width = 60
                DataGridView1.Columns(13).Width = 60
                DataGridView1.Columns(14).Width = 60
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


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(11).DefaultCellStyle.BackColor = Color.Bisque
                DataGridView1.Columns(12).DefaultCellStyle.BackColor = Color.LemonChiffon
                DataGridView1.Columns(14).DefaultCellStyle.BackColor = Color.DarkTurquoise
                DataGridView1.Columns(6).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(7).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(8).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(9).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(10).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(11).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(12).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(13).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(14).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(10).Width = 60
                DataGridView1.Columns(11).Width = 60
                DataGridView1.Columns(12).Width = 60
                DataGridView1.Columns(13).Width = 60
                DataGridView1.Columns(14).Width = 60
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar3()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "VENDEDOR" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(11).DefaultCellStyle.BackColor = Color.Bisque
                DataGridView1.Columns(12).DefaultCellStyle.BackColor = Color.LemonChiffon
                DataGridView1.Columns(14).DefaultCellStyle.BackColor = Color.DarkTurquoise
                DataGridView1.Columns(6).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(7).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(8).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(9).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(10).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(11).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(12).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(13).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(14).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(10).Width = 60
                DataGridView1.Columns(11).Width = 60
                DataGridView1.Columns(12).Width = 60
                DataGridView1.Columns(13).Width = 60
                DataGridView1.Columns(14).Width = 60
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim num1 As Integer
        num1 = e.RowIndex.ToString
        If e.ColumnIndex = 0 Then
            '' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'

            'num2 = e.ColumnIndex.ToString
            'Ingreso_observaciones.Show()
            'Ingreso_observaciones.TextBox4.Text = DataGridView1.Item(5, num1).Value()
            Detalle_Cotizacion.TextBox1.Text = 1
            Detalle_Cotizacion.TextBox2.Text = DataGridView1.Item(6, num1).Value()
            Detalle_Cotizacion.Label1.Text = "CODIGOS DE TELA"

            Detalle_Cotizacion.Show()
        End If
        If e.ColumnIndex = 1 Then
            Detalle_Cotizacion.TextBox1.Text = 2
            Detalle_Cotizacion.TextBox2.Text = DataGridView1.Item(6, num1).Value()
            Detalle_Cotizacion.Label1.Text = "AVIOS DE ACABADOS"
            Detalle_Cotizacion.Show()
        End If
        If e.ColumnIndex = 2 Then
            Detalle_Cotizacion.TextBox1.Text = 3
            Detalle_Cotizacion.TextBox2.Text = DataGridView1.Item(6, num1).Value()
            Detalle_Cotizacion.Label1.Text = "AVIOS DE COSTURA"
            Detalle_Cotizacion.Show()
        End If
        If e.ColumnIndex = 3 Then
            Detalle_Cotizacion.TextBox1.Text = 4
            Detalle_Cotizacion.TextBox2.Text = DataGridView1.Item(6, num1).Value()
            Detalle_Cotizacion.Label1.Text = "DETALLE DE LAVADOS"
            Detalle_Cotizacion.Show()
        End If
        If e.ColumnIndex = 4 Then
            Detalle_Cotizacion.TextBox1.Text = 5
            Detalle_Cotizacion.TextBox2.Text = DataGridView1.Item(6, num1).Value()
            Detalle_Cotizacion.Label1.Text = "DETALLE DE APLICACIONES"
            Detalle_Cotizacion.Show()
        End If
        If e.ColumnIndex = 5 Then
            Detalle_Cotizacion.TextBox1.Text = 6
            Detalle_Cotizacion.TextBox2.Text = DataGridView1.Item(6, num1).Value()
            Detalle_Cotizacion.Label1.Text = "MANO DE OBRA"
            Detalle_Cotizacion.Show()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar2()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscar3()
    End Sub
End Class