Public Class Reporte_conta
    Dim dt As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fg As New fventas
        Dim hh As New vventas

        Select Case ComboBox1.Text
            Case "ENERO" : hh.gmes = "01"
            Case "FEBRERO" : hh.gmes = "02"
            Case "MARZO" : hh.gmes = "03"
            Case "ABRIL" : hh.gmes = "04"
            Case "MAYO" : hh.gmes = "05"
            Case "JUNIO" : hh.gmes = "06"
            Case "JULIO" : hh.gmes = "07"
            Case "AGOSTO" : hh.gmes = "08"
            Case "SETIEMBRE" : hh.gmes = "09"
            Case "OCTUBRE" : hh.gmes = "10"
            Case "NOVIEMBRE" : hh.gmes = "11"
            Case "DICIEMBRE" : hh.gmes = "12"
        End Select
        hh.gANO = Label2.Text
        dt = fg.REPORTE_CONTA(hh)
        DataGridView1.DataSource = dt
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim jk As New Exportar
        jk.llenarExcel(DataGridView1)
    End Sub
    Private Sub buscar()
        Try

            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Cliente" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub Reporte_conta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = DateTime.Now.ToString("MM") - 1
    End Sub
End Class