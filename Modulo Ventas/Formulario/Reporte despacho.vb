Public Class Reporte_despacho
    Dim dt As New DataTable
    Private Sub Reporte_despacho_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim nn As New vcontablidad
        Dim hj As New fcontailidad
        ComboBox1.SelectedIndex = DateTimePicker1.Value.Month - 1

        Select Case ComboBox1.Text
            Case "ENERO" : nn.gMES = "01"
            Case "FEBRERO" : nn.gMES = "02"
            Case "MARZO" : nn.gMES = "03"
            Case "ABRIL" : nn.gMES = "04"
            Case "MAYO" : nn.gMES = "05"
            Case "JUNIO" : nn.gMES = "06"
            Case "JULIO" : nn.gMES = "07"
            Case "AGOSTO" : nn.gMES = "08"
            Case "SETIEMBRE" : nn.gMES = "09"
            Case "OCTUBRE" : nn.gMES = "10"
            Case "NOVIEMBRE" : nn.gMES = "11"
            Case "DICIEMBRE" : nn.gMES = "12"
        End Select

        nn.gperiodo = TextBox1.Text
        nn.gcia = Label5.Text
        dt = hj.REPORTE_HUARACHA2(nn)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(5).Width = 115
            DataGridView1.Columns(7).Width = 250
            DataGridView1.Columns(9).Visible = False

        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(5).Width = 115
                DataGridView1.Columns(7).Width = 250
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(11).Visible = False
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "FECHA" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(5).Width = 115
                DataGridView1.Columns(7).Width = 250
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(11).Visible = False
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim nn As New vcontablidad
        Dim hj As New fcontailidad
        Select Case ComboBox1.Text
            Case "ENERO" : nn.gMES = "01"
            Case "FEBRERO" : nn.gMES = "02"
            Case "MARZO" : nn.gMES = "03"
            Case "ABRIL" : nn.gMES = "04"
            Case "MAYO" : nn.gMES = "05"
            Case "JUNIO" : nn.gMES = "06"
            Case "JULIO" : nn.gMES = "07"
            Case "AGOSTO" : nn.gMES = "08"
            Case "SETIEMBRE" : nn.gMES = "09"
            Case "OCTUBRE" : nn.gMES = "10"
            Case "NOVIEMBRE" : nn.gMES = "11"
            Case "DICIEMBRE" : nn.gMES = "12"
        End Select
        nn.gcia = Label5.Text
        nn.gperiodo = TextBox1.Text
        dt = hj.REPORTE_HUARACHA2(nn)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(5).Width = 115
            DataGridView1.Columns(7).Width = 250
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(11).Visible = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim hl As New vcontablidad
        Dim hl1 As New vcontablidad
        Dim LL As New fcontailidad
        Dim k As Integer
        MsgBox("SE AGREGARA LAS  FACTURAS Y PROFORMAS")
        k = DataGridView1.Rows.Count
        DataGridView1.Columns(0).Visible = True
        For i = 0 To k - 1
            If Mid(Trim(DataGridView1.Rows(i).Cells(5).Value), 1, 2) = "14" Then
                hl.ggrm = Trim(DataGridView1.Rows(i).Cells(8).Value)
                hl.gcia = Label5.Text
                DataGridView1.Rows(i).Cells(2).Value = LL.buscar_factura_huaracha(hl)
            Else
                hl1.galmacen = Trim(DataGridView1.Rows(i).Cells(9).Value)
                hl1.gncom = Trim(DataGridView1.Rows(i).Cells(5).Value)
                hl1.gcia = Label5.Text
                DataGridView1.Rows(i).Cells(3).Value = LL.buscar_reportehh(hl1)
            End If


        Next

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim vg As New vcontablidad
            Dim vg1 As New fcontailidad
            Dim ju, JK As String
            Dim jñ As Date
            JK = ""
            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString


            If Mid(DataGridView1.Rows(num1).Cells(2).Value, 1, 1) = "F" Then
                JK = "01"
            Else

                If Mid(DataGridView1.Rows(num1).Cells(2).Value, 1, 1) = "B" Then
                    JK = "03"
                Else

                End If

            End If
            If Convert.ToString(DataGridView1.Rows(num1).Cells(2).Value) = "" Then

                Form_ventan.TextBox1.Text = Mid(DataGridView1.Rows(num1).Cells(3).Value, 1, 4)
                Form_ventan.TextBox2.Text = Mid(DataGridView1.Rows(num1).Cells(3).Value, 6, 8)
                Form_ventan.Show()
            Else
                vg.gfactura = DataGridView1.Rows(num1).Cells(2).Value
                vg.gcia = Label5.Text
                jñ = vg1.buscar_factura_fecha(vg)
                ju = Format(jñ, "dd/MM/yyyy")
                If Label5.Text = "01" Then
                    Try
                        Documento_Pdf.WebBrowser1.Navigate("\\172.21.0.1/AceptaService21/Peru_Prod/var/pdf/20508740361/2021/" & Mid(ju, 4, 2) & "/" & Mid(ju, 1, 2) & "/" & JK & "-" & Trim(DataGridView1.Rows(num1).Cells(2).Value) & ".pdf", False)
                        Documento_Pdf.Show()
                    Catch ex As Exception
                        MsgBox("NO SE ENCONTRO EL ARCHIVO")
                        Shell("\\172.21.0.1\AceptaService21\Peru_Prod\var\pdf\20508740361\2021", vbMaximizedFocus)
                    End Try
                Else
                    MsgBox("LOS PDF DE GRAUS NO SE GUARDAN EN EL SERVIDOR")
                End If



            End If


            'DataGridView1.Rows(num1).Cells(3).Value = True
            'DataGridView1.Columns(16).DefaultCellStyle.BackColor = Color.White
            'DataGridView1.Rows(num1).Cells(16).Value = ""
            'Button1.Visible = True
            'DataGridView1.Rows(num1).Cells(16).Selected = True
        Else
            If e.ColumnIndex = 1 Then
                Dim num1, num2 As Integer
                num1 = e.RowIndex.ToString
                num2 = e.ColumnIndex.ToString
                Reporte_Nia_Nsa.TextBox1.Text = Trim(DataGridView1.Rows(num1).Cells(11).Value)
                Reporte_Nia_Nsa.TextBox2.Text = 2
                Reporte_Nia_Nsa.TextBox3.Text = Trim(DataGridView1.Rows(num1).Cells(5).Value)
                Reporte_Nia_Nsa.TextBox4.Text = TextBox1.Text
                Reporte_Nia_Nsa.TextBox5.Text = Label5.Text
                Reporte_Nia_Nsa.Show()
            End If
        End If

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscar1()
    End Sub
End Class