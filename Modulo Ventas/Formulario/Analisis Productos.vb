Public Class Analisis_Productos
    Dim dt As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim jh As New fanalisis
        Dim hj As New vanalisis
        Dim HH As Double
        Dim l As Integer
        hj.gfini = DateTimePicker1.Value
        hj.gffin = DateTimePicker2.Value
        hj.gccia = Label5.Text
        dt = jh.buscar_informacion(hj)

        If dt.Rows.Count > 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(1).Width = 170
            DataGridView1.Columns(2).Width = 500
            DataGridView1.Columns(3).Width = 90
            DataGridView1.Columns(4).Width = 90
            l = DataGridView1.Rows.Count
            For i = 0 To l - 1
                HH = HH + DataGridView1.Rows(i).Cells(3).Value
            Next
            TextBox2.Text = HH
        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim HH As Double
            Dim l As Integer

            dv.RowFilter = "PRODUCTO" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(1).Width = 170
                DataGridView1.Columns(2).Width = 500
                DataGridView1.Columns(3).Width = 90
                DataGridView1.Columns(4).Width = 90
                l = DataGridView1.Rows.Count
                For i = 0 To l - 1
                    HH = HH + DataGridView1.Rows(i).Cells(3).Value
                Next
                TextBox2.Text = HH
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscarMes()
        Try
            Dim HH As Double
            Dim l As Integer
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "MES" & " like '%" & ComboBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(1).Width = 170
                DataGridView1.Columns(2).Width = 500
                DataGridView1.Columns(3).Width = 90
                DataGridView1.Columns(4).Width = 90
                l = DataGridView1.Rows.Count
                For i = 0 To l - 1
                    HH = HH + DataGridView1.Rows(i).Cells(3).Value
                Next
                TextBox2.Text = HH

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Analisis_Productos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub



    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        buscarMes()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim JL As New Exportar
        JL.llenarExcel(DataGridView1)
    End Sub
End Class