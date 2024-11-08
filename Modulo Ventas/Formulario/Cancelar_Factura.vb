Public Class Cancelar_Factura
    Dim dt As New DataTable
    Private Sub ACTUALIZAR()
        Dim gh As New ffactura
        Dim suma, suma1, suma2 As Double
        dt = gh.buscar_facturas
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(2).Width = 60
            DataGridView1.Columns(3).Width = 85
            DataGridView1.Columns(4).Width = 94
            DataGridView1.Columns(5).Width = 130
            DataGridView1.Columns(6).Width = 300
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(8).Width = 80
            DataGridView1.Columns(9).Width = 80
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).ReadOnly = True
            DataGridView1.Columns(4).ReadOnly = True
            DataGridView1.Columns(5).ReadOnly = True
            DataGridView1.Columns(6).ReadOnly = True
            DataGridView1.Columns(7).ReadOnly = True
            DataGridView1.Columns(8).ReadOnly = True
            DataGridView1.Columns(9).ReadOnly = True
            DataGridView1.Columns(10).ReadOnly = True
            DataGridView1.Columns(12).ReadOnly = True
            Dim i As Integer

            i = DataGridView1.Rows.Count
            For t = 0 To i - 1
                DataGridView1.Rows(t).Cells(11).ReadOnly = True

                DataGridView1.Rows(t).Cells(12).Value = DataGridView1.Rows(t).Cells(9).Value - DataGridView1.Rows(t).Cells(10).Value
                If DataGridView1.Rows(t).Cells(10).Value > "0.00" Then

                    DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.BurlyWood
                End If
                suma = suma + DataGridView1.Rows(t).Cells(9).Value
                suma1 = suma1 + DataGridView1.Rows(t).Cells(10).Value
                suma2 = suma2 + DataGridView1.Rows(t).Cells(12).Value
            Next
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            DataGridView1.Columns(4).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(8).Visible = False
            TextBox3.Text = Format(suma, "##,##00.00")
            TextBox4.Text = Format(suma1, "##,##00.00")
            TextBox5.Text = Format(suma2, "##,##00.00")
        End If
    End Sub
    Private Sub Cancelar_Factura_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ACTUALIZAR()
    End Sub
    Private Sub buscar()
        Try
            Dim suma, suma1, suma2 As Double
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(2).Width = 60
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 94
                DataGridView1.Columns(5).Width = 130
                DataGridView1.Columns(6).Width = 300
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(8).Width = 80
                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(2).ReadOnly = True
                DataGridView1.Columns(3).ReadOnly = True
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
                DataGridView1.Columns(7).ReadOnly = True
                DataGridView1.Columns(8).ReadOnly = True
                DataGridView1.Columns(9).ReadOnly = True
                DataGridView1.Columns(10).ReadOnly = True
                DataGridView1.Columns(12).ReadOnly = True
                Dim i As Integer

                i = DataGridView1.Rows.Count
                For t = 0 To i - 1
                    DataGridView1.Rows(t).Cells(11).ReadOnly = True

                    DataGridView1.Rows(t).Cells(12).Value = DataGridView1.Rows(t).Cells(9).Value - DataGridView1.Rows(t).Cells(10).Value
                    If DataGridView1.Rows(t).Cells(10).Value > "0.00" Then

                        DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.BurlyWood
                    End If
                    suma = suma + DataGridView1.Rows(t).Cells(7).Value
                    suma1 = suma1 + DataGridView1.Rows(t).Cells(10).Value
                    suma2 = suma2 + DataGridView1.Rows(t).Cells(12).Value
                Next




                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
                TextBox3.Text = Format(suma, "##,##00.00")
                TextBox4.Text = Format(suma1, "##,##00.00")
                TextBox5.Text = Format(suma2, "##,##00.00")
            Else


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


            dv.RowFilter = "FACTURA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(2).Width = 60
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 94
                DataGridView1.Columns(5).Width = 130
                DataGridView1.Columns(6).Width = 300
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(8).Width = 80
                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(2).ReadOnly = True
                DataGridView1.Columns(3).ReadOnly = True
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
                DataGridView1.Columns(7).ReadOnly = True
                DataGridView1.Columns(8).ReadOnly = True
                DataGridView1.Columns(9).ReadOnly = True
                DataGridView1.Columns(10).ReadOnly = True
                DataGridView1.Columns(12).ReadOnly = True
                Dim i As Integer
                Dim suma, suma1, suma2 As Double
                i = DataGridView1.Rows.Count
                For t = 0 To i - 1
                    DataGridView1.Rows(t).Cells(11).ReadOnly = True

                    DataGridView1.Rows(t).Cells(12).Value = DataGridView1.Rows(t).Cells(9).Value - DataGridView1.Rows(t).Cells(10).Value
                    If DataGridView1.Rows(t).Cells(10).Value > "0.00" Then

                        DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.BurlyWood
                    End If
                    suma = suma + DataGridView1.Rows(t).Cells(7).Value
                    suma1 = suma1 + DataGridView1.Rows(t).Cells(10).Value
                    suma2 = suma2 + DataGridView1.Rows(t).Cells(12).Value
                Next




                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
                TextBox3.Text = Format(suma, "##,##00.00")
                TextBox4.Text = Format(suma1, "##,##00.00")
                TextBox5.Text = Format(suma2, "##,##00.00")
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar1()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a As Integer
        Dim fh As New ffactura
        Dim fh1 As New vfactura
        Dim jk As New vdet_regis_fact
        i = DataGridView1.Rows.Count
        a = 0
        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                fh1.gsfactu = Microsoft.VisualBasic.Left(DataGridView1.Rows(a).Cells(5).Value, 4)
                fh1.gnfactu = Microsoft.VisualBasic.Right(DataGridView1.Rows(a).Cells(5).Value, 8)
                fh1.gperiodo = DataGridView1.Rows(a).Cells(2).Value
                fh.actualizar_factura(fh1)
                If DataGridView1.Rows(a).Cells(10).Value > "0.00" Then

                    DataGridView1.Rows(a).DefaultCellStyle.BackColor = Color.BurlyWood
                End If
            End If
            If Me.DataGridView1.Rows(a).Cells(1).Value = True Then
                    fh1.gsfactu = Microsoft.VisualBasic.Left(DataGridView1.Rows(a).Cells(5).Value, 4)
                    fh1.gnfactu = Microsoft.VisualBasic.Right(DataGridView1.Rows(a).Cells(5).Value, 8)
                    fh1.gperiodo = DataGridView1.Rows(a).Cells(2).Value
                    fh1.gMONTO = DataGridView1.Rows(a).Cells(10).Value + DataGridView1.Rows(a).Cells(11).Value
                fh.actualizar_factura_MONTO(fh1)
                jk.gperiodo = DataGridView1.Rows(a).Cells(2).Value
                jk.gcomprobante = DataGridView1.Rows(a).Cells(5).Value
                jk.gcliente = DataGridView1.Rows(a).Cells(6).Value
                jk.gvalor_venta = DataGridView1.Rows(a).Cells(7).Value
                jk.gigv = DataGridView1.Rows(a).Cells(8).Value
                jk.gtotal = DataGridView1.Rows(a).Cells(9).Value
                jk.gpagado = DataGridView1.Rows(a).Cells(10).Value
                jk.gparcial = DataGridView1.Rows(a).Cells(11).Value
                jk.gobservacion = DataGridView1.Rows(a).Cells(13).Value
                jk.gfecha = DateTimePicker1.Value
                fh.insertar_ventas_fac_reg(jk)
            End If

        Next
        MsgBox("SE CANCELO LAS FACTURAS SOLICITADAS")
        ACTUALIZAR()
        TextBox2.Text = ""
        TextBox1.Text = ""
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            If DataGridView1.Rows(num1).Cells(1).Value = True Then
                DataGridView1.Rows(num1).Cells(1).Value = False
                DataGridView1.Columns(11).DefaultCellStyle.BackColor = Color.White
                DataGridView1.Rows(num1).DefaultCellStyle.BackColor = Color.White

            Else
                If DataGridView1.Rows(num1).Cells(1).Value = False Then
                    DataGridView1.Rows(num1).Cells(1).Value = True
                    DataGridView1.Columns(11).DefaultCellStyle.BackColor = Color.DarkTurquoise
                    DataGridView1.Rows(num1).Cells(11).Selected = True
                    DataGridView1.Rows(num1).Cells(11).Value = DataGridView1.Rows(num1).Cells(9).Value
                    DataGridView1.Rows(num1).DefaultCellStyle.BackColor = Color.DarkTurquoise
                End If

            End If

        End If
        If e.ColumnIndex = 1 Then

                ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
                Dim num11, num12 As Integer
                num11 = e.RowIndex.ToString
                num12 = e.ColumnIndex.ToString
                DataGridView1.Rows(num11).Cells(0).Value = False
                DataGridView1.Rows(num11).Cells(11).ReadOnly = False
                DataGridView1.Columns(11).DefaultCellStyle.BackColor = Color.DarkTurquoise
                DataGridView1.Rows(num11).Cells(11).Selected = True
            DataGridView1.Rows(num11).DefaultCellStyle.BackColor = Color.DarkTurquoise
        End If


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 13 Then
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            Observacion2.TextBox2.Text = num1

            Observacion2.TextBox1.Text = DataGridView1.Rows(num1).Cells(13).Value
            Observacion2.Show()
        End If
        If e.RowIndex < 0 Then
            Dim i As Integer
            Dim suma, suma1, suma2 As Double
            i = DataGridView1.Rows.Count
            For t = 0 To i - 1
                DataGridView1.Rows(t).Cells(11).ReadOnly = True

                DataGridView1.Rows(t).Cells(12).Value = DataGridView1.Rows(t).Cells(9).Value - DataGridView1.Rows(t).Cells(10).Value
                If DataGridView1.Rows(t).Cells(10).Value > "0.00" Then

                    DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.BurlyWood
                End If
                suma = suma + DataGridView1.Rows(t).Cells(9).Value
                suma1 = suma1 + DataGridView1.Rows(t).Cells(10).Value
                suma2 = suma2 + DataGridView1.Rows(t).Cells(12).Value
            Next
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            DataGridView1.Columns(4).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            TextBox3.Text = Format(suma, "##,##00.00")
            TextBox4.Text = Format(suma1, "##,##00.00")
            TextBox5.Text = Format(suma2, "##,##00.00")
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex < 0 Then
            Dim i As Integer
            Dim suma, suma1, suma2 As Double
            i = DataGridView1.Rows.Count
            For t = 0 To i - 1
                DataGridView1.Rows(t).Cells(11).ReadOnly = True

                DataGridView1.Rows(t).Cells(12).Value = DataGridView1.Rows(t).Cells(9).Value - DataGridView1.Rows(t).Cells(10).Value
                If DataGridView1.Rows(t).Cells(10).Value > "0.00" Then

                    DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.BurlyWood
                End If
                suma = suma + DataGridView1.Rows(t).Cells(9).Value
                suma1 = suma1 + DataGridView1.Rows(t).Cells(10).Value
                suma2 = suma2 + DataGridView1.Rows(t).Cells(12).Value
            Next
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            DataGridView1.Columns(4).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            TextBox3.Text = Format(suma, "##,##00.00")
            TextBox4.Text = Format(suma1, "##,##00.00")
            TextBox5.Text = Format(suma2, "##,##00.00")
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        Dim i As Integer
        Dim suma, suma1, suma2 As Double
        i = DataGridView1.Rows.Count
        For t = 0 To i - 1
            DataGridView1.Rows(t).Cells(11).ReadOnly = True

            DataGridView1.Rows(t).Cells(12).Value = DataGridView1.Rows(t).Cells(9).Value - DataGridView1.Rows(t).Cells(10).Value
            If DataGridView1.Rows(t).Cells(10).Value > "0.00" Then

                DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.BurlyWood
            End If
            suma = suma + DataGridView1.Rows(t).Cells(9).Value
            suma1 = suma1 + DataGridView1.Rows(t).Cells(10).Value
            suma2 = suma2 + DataGridView1.Rows(t).Cells(12).Value
        Next
        DataGridView1.Columns(1).Frozen = True
        DataGridView1.Columns(2).Frozen = True
        DataGridView1.Columns(3).Frozen = True
        DataGridView1.Columns(4).Frozen = True
        DataGridView1.Columns(5).Frozen = True
        DataGridView1.Columns(5).Frozen = True
        TextBox3.Text = Format(suma, "##,##00.00")
        TextBox4.Text = Format(suma1, "##,##00.00")
        TextBox5.Text = Format(suma2, "##,##00.00")
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
End Class