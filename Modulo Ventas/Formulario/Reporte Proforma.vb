Public Class Reporte_Proforma
    Dim HG As New DataTable
    Private Sub Reporte_Proforma_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim jh As New fcontailidad
        Dim kk As New vcontablidad

        ComboBox1.SelectedIndex = DateTimePicker1.Value.Month - 1

        Select Case ComboBox1.Text
            Case "ENERO" : kk.gMES = "01"
            Case "FEBRERO" : kk.gMES = "02"
            Case "MARZO" : kk.gMES = "03"
            Case "ABRIL" : kk.gMES = "04"
            Case "MAYO" : kk.gMES = "05"
            Case "JUNIO" : kk.gMES = "06"
            Case "JULIO" : kk.gMES = "07"
            Case "AGOSTO" : kk.gMES = "08"
            Case "SETIEMBRE" : kk.gMES = "09"
            Case "OCTUBRE" : kk.gMES = "10"
            Case "NOVIEMBRE" : kk.gMES = "11"
            Case "DICIEMBRE" : kk.gMES = "12"
        End Select

        kk.gperiodo = TextBox2.Text
        kk.gcia = Label4.Text
        HG = jh.REPORTE_HUARACHA(kk)

        If HG.Rows.Count <> 0 Then
            DataGridView1.DataSource = HG
            DataGridView1.Columns(4).Width = 200
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(5).Width = 80
            DataGridView1.Columns(6).Width = 80

        Else
            DataGridView1.DataSource = ""
        End If
    End Sub
    Private Sub buscar()
        Try
            Dim DT As New DataTable
            DataGridView1.DataSource = DT
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(4).Width = 200
                DataGridView1.Columns(1).Width = 80
                DataGridView1.Columns(5).Width = 80
                DataGridView1.Columns(6).Width = 80

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim jh As New fcontailidad
        Dim kk As New vcontablidad

        Select Case ComboBox1.Text
            Case "ENERO" : kk.gMES = "01"
            Case "FEBRERO" : kk.gMES = "02"
            Case "MARZO" : kk.gMES = "03"
            Case "ABRIL" : kk.gMES = "04"
            Case "MAYO" : kk.gMES = "05"
            Case "JUNIO" : kk.gMES = "06"
            Case "JULIO" : kk.gMES = "07"
            Case "AGOSTO" : kk.gMES = "08"
            Case "SETIEMBRE" : kk.gMES = "09"
            Case "OCTUBRE" : kk.gMES = "10"
            Case "NOVIEMBRE" : kk.gMES = "11"
            Case "DICIEMBRE" : kk.gMES = "12"
        End Select
        kk.gcia = Label4.Text
        kk.gperiodo = TextBox2.Text
        HG = jh.REPORTE_HUARACHA(kk)

        If HG.Rows.Count <> 0 Then
            DataGridView1.DataSource = HG
            DataGridView1.Columns(4).Width = 200
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(5).Width = 80
            DataGridView1.Columns(6).Width = 80
        Else
            DataGridView1.DataSource = ""
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        If e.ColumnIndex = 0 Then

            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("DESEA IMPRIMIR EL REGISTRO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Dim num1, num2 As Integer
                num1 = e.RowIndex.ToString
                num2 = e.ColumnIndex.ToString
                Form_ventan.TextBox1.Text = Mid(DataGridView1.Rows(num1).Cells(2).Value, 1, 4)
                Form_ventan.TextBox2.Text = Mid(DataGridView1.Rows(num1).Cells(2).Value, 6, 8)
                Form_ventan.Show()
            End If

        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class