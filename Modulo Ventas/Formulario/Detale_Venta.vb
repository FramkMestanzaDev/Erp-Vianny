Public Class Detale_Venta
    Dim dt As New DataTable
    Private Sub Detale_Venta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fg As New fregistro_venta
        Dim fg1 As New vregistroventa

        fg1.gncom_3 = TextBox1.Text
        fg1.gcperiodo_3 = TextBox2.Text
        fg1.gempresa = Label1.Text
        dt = fg.mostrar_detalle_venta(fg1)
        DataGridView1.DataSource = dt
        DataGridView1.Columns(0).Width = 80
        DataGridView1.Columns(1).Width = 140
        DataGridView1.Columns(2).Width = 250
        DataGridView1.Columns(0).Frozen = True
        DataGridView1.Columns(1).Frozen = True
        DataGridView1.Columns(2).Frozen = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 7 Then
            Dim respuesta As DialogResult

            respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                Dim vg As New vcontablidad
                Dim vg1 As New fcontailidad

                Dim num1, num2 As Integer
                num1 = e.RowIndex.ToString
                num2 = e.ColumnIndex.ToString
                Reporte_Nia_Nsa.TextBox1.Text = Mid(Trim(DataGridView1.Rows(num1).Cells(7).Value), 1, 2)
                Reporte_Nia_Nsa.TextBox2.Text = 2
                Reporte_Nia_Nsa.TextBox3.Text = Mid(Trim(DataGridView1.Rows(num1).Cells(7).Value), 3, 8)
                Reporte_Nia_Nsa.TextBox4.Text = TextBox2.Text
                Reporte_Nia_Nsa.TextBox5.Text = Label1.Text
                Reporte_Nia_Nsa.Show()
            End If

        End If

    End Sub
End Class