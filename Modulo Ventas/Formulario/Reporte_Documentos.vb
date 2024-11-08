Public Class Reporte_Documentos
    Private Sub Reporte_Documentos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Chart1_Click(sender As Object, e As EventArgs) Handles Chart1.Click

    End Sub
    Dim dt As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim fg As New freportedoc
        Dim fg1 As New vreportedoc
        Dim valo As String
        valo = UCase(MonthName(Month(DateTimePicker1.Value)))
        fg1.gfechaini = DateSerial(Year(DateTimePicker1.Value), Month(DateTimePicker1.Value), 1)
        fg1.gfechafin = DateSerial(Year(DateTimePicker1.Value), Month(DateTimePicker1.Value) + 1, 0)
        dt = fg.reportedoc1(fg1)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            Me.Chart1.Series.Clear()
            Me.Chart1.Series.Add(valo)
            Me.Chart1.Series.Item(0).BorderColor = Color.Aqua
            Dim ad As Integer
            ad = DataGridView1.Rows.Count
            For i = 0 To ad - 1
                Me.Chart1.Series(valo).Points.AddXY(DataGridView1.Rows(i).Cells(0).Value, DataGridView1.Rows(i).Cells(1).Value)
            Next

            'Me.Chart1.Series("JULIO").Points.AddXY("Y.CARRILLO", 40)
            'Me.Chart1.Series("JULIO").Points.AddXY("H.RIVERA", 55)
            'Me.Chart1.Series("JULIO").Points.AddXY("M.VALDEZ", 60)
            'Me.Chart1.Series("JULIO").Points.AddXY("Y.PAZ", 70)
        Else
            MsgBox("NO EXISTE INFORMACION PARA MOSTRAR", vbExclamation)
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Chart1.Series.Clear()
        Dim fg As New freportedoc
        Dim fg1 As New vreportedoc
        Dim j, KH, KH1, KH2, KH3, KH4, KH5, KH6, KH7, KH8, KH9, KH10, KH11, KH12, KH13, KH14 As Integer
        fg1.gfechaini = DateSerial(Year(DateTimePicker1.Value), Month(DateTimePicker1.Value) - 2, 1)
        fg1.gfechafin = DateSerial(Year(DateTimePicker1.Value), Month(DateTimePicker1.Value) + 1, 0)
        dt = fg.reportedoc2(fg1)

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
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            Me.Chart1.Series.Add(UCase(MonthName(Month(DateTimePicker1.Value))))
            Me.Chart1.Series.Add(UCase(MonthName(Month(DateTimePicker1.Value) - 1)))
            Me.Chart1.Series.Add(UCase(MonthName(Month(DateTimePicker1.Value) - 2)))
            j = DataGridView1.Rows.Count
            'Me.Chart1.ChartAreas(0).Area3DStyle.Enable3D = True
            For i = 0 To j - 1

                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "MLAURENTE" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value))) Then

                    KH = KH + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "MLAURENTE" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 1)) Then

                    KH1 = KH1 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "MLAURENTE" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 2)) Then

                    KH2 = KH2 + 1
                End If
                'FEWFWEEEEEEEEEEEEEEEEEEEEEEEEEEEEE


                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "MVALDEZ" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value))) Then

                    KH3 = KH3 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "MVALDEZ" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 1)) Then

                    KH4 = KH4 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "MVALDEZ" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 2)) Then

                    KH5 = KH5 + 1
                End If
                '-----------------------------------------------------------------------

                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "YPAZ" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value))) Then
                    KH6 = KH6 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "YPAZ" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 1)) Then

                    KH7 = KH7 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "YPAZ" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 2)) Then

                    KH8 = KH8 + 1
                End If

                '------------------------------------------------
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "YCARRILLO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value))) Then
                    KH9 = KH9 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "YCARRILLO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 1)) Then

                    KH10 = KH10 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "YCARRILLO" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 2)) Then

                    KH11 = KH11 + 1
                End If
                '------------------------------------------------
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "HRIVERA" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value))) Then
                    KH12 = KH12 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "HRIVERA" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 1)) Then

                    KH13 = KH13 + 1
                End If
                If DataGridView1.Rows(i).Cells(0).Value.ToString.Trim = "HRIVERA" And UCase(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) = UCase(MonthName(Month(DateTimePicker1.Value) - 2)) Then

                    KH14 = KH14 + 1
                End If
            Next

            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value)))).Points.AddXY("MLAURENTE", KH)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 1))).Points.AddXY("MLAURENTE", KH1)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 2))).Points.AddXY("MLAURENTE", KH2)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value)))).Points.AddXY("MVALDEZ", KH3)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 1))).Points.AddXY("MVALDEZ", KH4)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 2))).Points.AddXY("MVALDEZ", KH5)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value)))).Points.AddXY("YPAZ", KH6)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 1))).Points.AddXY("YPAZ", KH7)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 2))).Points.AddXY("YPAZ", KH8)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value)))).Points.AddXY("YCARRILLO", KH9)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 1))).Points.AddXY("YCARRILLO", KH10)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 2))).Points.AddXY("YCARRILLO", KH11)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value)))).Points.AddXY("HRIVERA", KH12)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 1))).Points.AddXY("HRIVERA", KH13)
            Me.Chart1.Series(UCase(MonthName(Month(DateTimePicker1.Value) - 2))).Points.AddXY("HRIVERA", KH14)

        End If

    End Sub
End Class