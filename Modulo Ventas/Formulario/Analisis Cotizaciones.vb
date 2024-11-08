Imports System.Windows.Forms.DataVisualization.Charting
Public Class Analisis_Cotizaciones
    Dim DT, dt2, DT3 As New DataTable

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim fg As New FCOTIZACION
        Dim fg1 As New VCOTIZACION
        Dim valo As String
        valo = UCase(MonthName(Month(DateTimePicker1.Value)))
        fg1.gfinicial = DateSerial(Year(DateTimePicker1.Value), Month(DateTimePicker1.Value), 1)
        fg1.gffinal = DateSerial(Year(DateTimePicker1.Value), Month(DateTimePicker1.Value) + 1, 0)
        DT = fg.reportedoc1(fg1)
        If DT.Rows.Count <> 0 Then
            DataGridView1.DataSource = DT
            Me.Chart1.Series.Clear()
            Me.Chart1.ChartAreas.Clear()
            Me.Chart1.Series.Add(valo)
            Me.Chart1.ChartAreas.Add(valo)
            Me.Chart1.Series.Item(0).BorderColor = Color.Aqua
            Dim ad As Integer
            ad = DataGridView1.Rows.Count
            For i = 0 To ad - 1
                Me.Chart1.Series(valo).Points.AddXY(DataGridView1.Rows(i).Cells(0).Value, DataGridView1.Rows(i).Cells(1).Value)

            Next

            Me.Chart1.Series(valo).IsValueShownAsLabel = True
            Me.Chart1.Series(valo)("DrawingStyle") = "Cylinder"
            Me.Chart1.Series(valo)("LabelStyle") = "Bottom"
            Me.Chart1.Series(valo).ChartType = SeriesChartType.Column
            'Me.Chart1.Series(valo).ChartType = SeriesChartType.Line
            'Me.Chart1.ChartAreas(0).Area3DStyle.Enable3D = True
        Else
            MsgBox("NO EXISTE INFORMACION PARA MOSTRAR", vbExclamation)
        End If
    End Sub

    Private Sub Analisis_Cotizaciones_Load(sender As Object, e As EventArgs)

    End Sub

    Private Sub Analisis_Cotizaciones_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim fg As New FCOTIZACION
        Dim fg1 As New VCOTIZACION

        fg1.gvendedor = ComboBox1.Text.ToString.Trim
        DT3 = fg.ganancias(fg1)

        If DT3.Rows.Count <> 0 Then
            Me.Chart1.Series.Clear()
            DataGridView1.DataSource = DT3
            Me.Chart1.Series.Add(ComboBox1.Text)
            Dim ad As Integer
            ad = DataGridView1.Rows.Count
            For i = 0 To ad - 1
                Me.Chart1.Series(0).Points.AddXY(Convert.ToInt32(DataGridView1.Rows(i).Cells(0).Value), Convert.ToDouble(DataGridView1.Rows(i).Cells(1).Value))

            Next
            Me.Chart1.Series(0).IsValueShownAsLabel = True
            Me.Chart1.Series(0)("DrawingStyle") = "Cylinder"
            Me.Chart1.Series(0)("LabelStyle") = "Bottom"

            Me.Chart1.ChartAreas(0).Area3DStyle.Enable3D = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Chart1.Series.Clear()
        Me.Chart1.ChartAreas.Clear()
        Dim fg As New FCOTIZACION
        Dim fg1 As New VCOTIZACION
        Dim j, KH, KH1, KH2, KH3, KH4, KH5, KH6, KH7, KH8, KH9, KH10, KH11, KH12, KH13, KH14 As Integer
        fg1.gfinicial = DateSerial(Year(DateTimePicker1.Value), Month(DateTimePicker1.Value) - 2, 1)
        fg1.gffinal = DateSerial(Year(DateTimePicker1.Value), Month(DateTimePicker1.Value) + 1, 0)
        dt2 = fg.reportedoc2(fg1)

        KH = 0
        KH1 = 0
        KH2 = 0
        KH3 = 0
        KH4 = 0
        KH5 = 0
        KH6 = 0
        KH8 = 0
        KH7 = 0
        KH9 = 0
        KH10 = 0
        KH11 = 0
        KH12 = 0
        KH13 = 0
        KH14 = 0
        If dt2.Rows.Count <> 0 Then

            DataGridView1.DataSource = dt2
            Me.Chart1.Series.Add(UCase(MonthName(Month(DateTimePicker1.Value))))
            Me.Chart1.Series.Add(UCase(MonthName(Month(DateTimePicker1.Value) - 1)))
            Me.Chart1.Series.Add(UCase(MonthName(Month(DateTimePicker1.Value) - 2)))
            Me.Chart1.ChartAreas.Add(UCase(MonthName(Month(DateTimePicker1.Value))))
            Me.Chart1.Series(0).IsValueShownAsLabel = True
            Me.Chart1.Series(1).IsValueShownAsLabel = True
            Me.Chart1.Series(2).IsValueShownAsLabel = True
            Me.Chart1.Series(0)("DrawingStyle") = "Cylinder"
            Me.Chart1.Series(1)("DrawingStyle") = "Emboss"
            Me.Chart1.Series(2)("DrawingStyle") = "Wedge"
            Me.Chart1.Series(0)("LabelStyle") = "Bottom"
            Me.Chart1.Series(1)("LabelStyle") = "Bottom"
            Me.Chart1.Series(2)("LabelStyle") = "Bottom"
            j = DataGridView1.Rows.Count
            For i = 0 To j - 1

                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "VPIZARRO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value))) Then

                    KH = KH + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "VPIZARRO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 1)) Then

                    KH1 = KH1 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "VPIZARRO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 2)) Then

                    KH2 = KH2 + 1
                End If
                'FEWFWEEEEEEEEEEEEEEEEEEEEEEEEEEEEE


                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "GBALVIN" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value))) Then

                    KH3 = KH3 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "GBALVIN" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 1)) Then

                    KH4 = KH4 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "GBALVIN" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 2)) Then

                    KH5 = KH5 + 1
                End If
                '-----------------------------------------------------------------------

                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "GBEDON" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value))) Then
                    KH6 = KH6 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "GBEDON" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 1)) Then

                    KH7 = KH7 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "GBEDON" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 2)) Then

                    KH8 = KH8 + 1
                End If

                '------------------------------------------------
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "DBRAVO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value))) Then
                    KH9 = KH9 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "DBRAVO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 1)) Then

                    KH10 = KH10 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "DBRAVO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 2)) Then

                    KH11 = KH11 + 1
                End If
                '------------------------------------------------
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "VINCIO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value))) Then
                    KH12 = KH12 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "VINCIO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 1)) Then

                    KH13 = KH13 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "VINCIO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 2)) Then

                    KH14 = KH14 + 1
                End If
            Next

            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value)))).Points.AddXY("VPIZARRO", KH)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 1))).Points.AddXY("VPIZARRO", KH1)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 2))).Points.AddXY("VPIZARRO", KH2)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value)))).Points.AddXY("GBALVIN", KH3)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 1))).Points.AddXY("GBALVIN", KH4)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 2))).Points.AddXY("GBALVIN", KH5)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value)))).Points.AddXY("GBEDON", KH6)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 1))).Points.AddXY("GBEDON", KH7)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 2))).Points.AddXY("GBEDON", KH8)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value)))).Points.AddXY("DBRAVO", KH9)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 1))).Points.AddXY("DBRAVO", KH10)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 2))).Points.AddXY("DBRAVO", KH11)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value)))).Points.AddXY("VINCIO", KH12)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 1))).Points.AddXY("VINCIO", KH13)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 2))).Points.AddXY("VINCIO", KH14)

        End If

    End Sub
End Class