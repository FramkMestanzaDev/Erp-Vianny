Public Class Orden_Costura
    Dim dt As New DataTable

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            actualizar()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            If DataGridView1.Rows(num1).Cells(11).Value.ToString.Trim = "03" Then
                DataGridView1.Rows(num1).Cells(0).ReadOnly = True
            Else
                DataGridView1.Rows(num1).Cells(0).Value = True
                DataGridView1.Rows(num1).Cells(1).Value = True
                'DataGridView1.Columns(15).DefaultCellStyle.BackColor = Color.White
                DataGridView1.Rows(num1).Cells(9).Value = ""
                DataGridView1.Rows(num1).Cells(9).ReadOnly = False


                DataGridView1.Rows(num1).Cells(9).Selected = True
            End If


        Else
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            If e.ColumnIndex = 1 Then

                If DataGridView1.Rows(num1).Cells(11).Value.ToString.Trim = "03" Then
                    DataGridView1.Rows(num1).Cells(1).ReadOnly = True
                Else

                    DataGridView1.Rows(num1).Cells(1).Value = True
                    DataGridView1.Rows(num1).Cells(0).Value = False
                    DataGridView1.Rows(num1).Cells(9).ReadOnly = False


                    DataGridView1.Rows(num1).Cells(9).Selected = True
                End If


            End If
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "C_PARCIAL" Then
            Dim va, va1 As Double
            va = DataGridView1.Rows(e.RowIndex).Cells(9).Value
            va1 = Format(va, "##,##00.00")

            If va1 > DataGridView1.Rows(e.RowIndex).Cells(8).Value Then
                DataGridView1.Rows(e.RowIndex).Cells(9).Value = 0.00

                MsgBox("LA CANTIDAD INGRESADA EXEDE A LO QUE SE LE ENTREGO AL TALLER")
            Else
                If va1 = DataGridView1.Rows(e.RowIndex).Cells(8).Value Then
                    DataGridView1.Rows(e.RowIndex).Cells(0).Value = True
                    MsgBox("CON LA CANTIDAD INGRESADA SE CANCELARA EL CORTE")
                Else
                    DataGridView1.Rows(e.RowIndex).Cells(0).Value = False
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim hj As New vorden_costura
        Dim hj3 As New vorden_costura
        Dim hj1 As New forden_costura
        Dim Ia As Integer
        Ia = DataGridView1.Rows.Count
        For a = 0 To Ia - 1
            If Me.DataGridView1.Rows(a).Cells(1).Value = True Then

                hj.gcantidad = DataGridView1.Rows(a).Cells(9).Value + DataGridView1.Rows(a).Cells(7).Value
                hj.gorden = DataGridView1.Rows(a).Cells(3).Value.ToString.Trim
                hj1.actualizar_cantidad(hj)
                hj3.gorden_c = DataGridView1.Rows(a).Cells(3).Value
                hj3.gop_oc = DataGridView1.Rows(a).Cells(4).Value
                hj3.gocorte = DataGridView1.Rows(a).Cells(5).Value
                hj3.gprendasc = DataGridView1.Rows(a).Cells(6).Value
                hj3.gadelanto = DataGridView1.Rows(a).Cells(9).Value
                hj3.gruc = DataGridView1.Rows(a).Cells(10).Value
                hj3.gfecha = DateTimePicker1.Value
                hj1.insertra_costura(hj3)

            End If
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then


                hj.gorden = DataGridView1.Rows(a).Cells(3).Value.ToString.Trim
                hj1.actualizar_costura(hj)


            End If
        Next

        MsgBox("SE ACTUALIZO LO SELECCIONADO")
        actualizar()

    End Sub
    Sub actualizar()
        Dim fo As New forden_costura
        Dim fo1 As New vorden_costura
        Dim Ia As Integer
        fo1.gop = TextBox1.Text
        dt = fo.buscar_orden_costuraop(fo1)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(2).Width = 220
            DataGridView1.Columns(3).Width = 80
            DataGridView1.Columns(4).Width = 90
            DataGridView1.Columns(5).Width = 70
            DataGridView1.Columns(6).Width = 90
            DataGridView1.Columns(7).Width = 90
            DataGridView1.Columns(8).Width = 90
            DataGridView1.Columns(12).Width = 90
            DataGridView1.Columns(6).ReadOnly = True
            DataGridView1.Columns(7).ReadOnly = True
            DataGridView1.Columns(8).ReadOnly = True
            DataGridView1.Columns(9).ReadOnly = True
            Ia = DataGridView1.Rows.Count
            For I = 0 To Ia - 1
                DataGridView1.Rows(I).Cells(8).Value = DataGridView1.Rows(I).Cells(6).Value - DataGridView1.Rows(I).Cells(7).Value

                If DataGridView1.Rows(I).Cells(11).Value.ToString.Trim = "03" Then

                    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.DarkTurquoise
                    DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White
                    DataGridView1.Rows(I).Cells(0).ReadOnly = True
                    DataGridView1.Rows(I).Cells(1).ReadOnly = True
                    DataGridView1.Rows(I).Cells(0).Value = False
                    DataGridView1.Rows(I).Cells(1).Value = False
                    DataGridView1.Rows(I).Cells(12).Value = "TERMINADO"
                Else
                    DataGridView1.Rows(I).Cells(12).Value = "PENDIENTE"
                End If

            Next
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(7).DefaultCellStyle.BackColor = Color.Khaki
            DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.Khaki
        Else
            MsgBox("LA OP INGRESADA NO EXISTE")
        End If
    End Sub

    Private Sub Orden_Costura_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class