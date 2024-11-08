Public Class Error_ventas
    Private Sub Error_ventas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = DateTime.Now.ToString("MM") - 1
    End Sub
    Dim DT As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim gh As New fventas
        Dim gh1 As New vventas
        Dim j As Integer
        Select Case ComboBox1.Text
            Case "ENERO" : gh1.gmes = "01"
            Case "FEBRERO" : gh1.gmes = "02"
            Case "MARZO" : gh1.gmes = "03"
            Case "ABRIL" : gh1.gmes = "04"
            Case "MAYO" : gh1.gmes = "05"
            Case "JUNIO" : gh1.gmes = "06"
            Case "JULIO" : gh1.gmes = "07"
            Case "AGOSTO" : gh1.gmes = "08"
            Case "SETIEMBRE" : gh1.gmes = "09"
            Case "OCTUBRE" : gh1.gmes = "10"
            Case "NOVIEMBRE" : gh1.gmes = "11"
            Case "DICIEMBRE" : gh1.gmes = "12"
        End Select
        gh1.gANO = Label4.Text
        gh1.gccia = Label5.Text
        DT = gh.ERROR_VENTAS(gh1)

        If DT.Rows.Count <> 0 Then
            DataGridView1.DataSource = DT
            j = DataGridView1.Rows.Count

            For i = 0 To j - 1
                DataGridView1.Columns(0).Width = 85
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(2).Width = 220
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 85
                DataGridView1.Columns(5).Width = 65
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 150
                DataGridView1.Columns(8).Width = 270

                DataGridView1.Columns(3).DefaultCellStyle.BackColor = Color.DarkGray
                DataGridView1.Columns(3).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Columns(4).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Columns(4).DefaultCellStyle.ForeColor = Color.White

                DataGridView1.Columns(7).DefaultCellStyle.BackColor = Color.DarkGray
                DataGridView1.Columns(7).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Columns(8).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Columns(8).DefaultCellStyle.ForeColor = Color.White
            Next
        Else
            MsgBox("NO HAY INFORMACION REGISTRADA EL MES SELECCIONADO")
            DataGridView1.DataSource = ""
        End If

    End Sub
    Private Sub buscar()
        Try

            DataGridView1.DataSource = DT
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim j As Integer

            dv.RowFilter = "VENDEDOR_F" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                j = DataGridView1.Rows.Count

                For i = 0 To j - 1
                    DataGridView1.Columns(0).Width = 85
                    DataGridView1.Columns(1).Width = 90
                    DataGridView1.Columns(2).Width = 220
                    DataGridView1.Columns(3).Width = 85
                    DataGridView1.Columns(4).Width = 85
                    DataGridView1.Columns(5).Width = 65
                    DataGridView1.Columns(6).Width = 80
                    DataGridView1.Columns(7).Width = 150
                    DataGridView1.Columns(8).Width = 270

                    DataGridView1.Columns(3).DefaultCellStyle.BackColor = Color.DarkGray
                    DataGridView1.Columns(3).DefaultCellStyle.ForeColor = Color.White
                    DataGridView1.Columns(4).DefaultCellStyle.BackColor = Color.DarkKhaki
                    DataGridView1.Columns(4).DefaultCellStyle.ForeColor = Color.White

                    DataGridView1.Columns(7).DefaultCellStyle.BackColor = Color.DarkGray
                    DataGridView1.Columns(7).DefaultCellStyle.ForeColor = Color.White
                    DataGridView1.Columns(8).DefaultCellStyle.BackColor = Color.DarkKhaki
                    DataGridView1.Columns(8).DefaultCellStyle.ForeColor = Color.White
                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class