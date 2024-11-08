Public Class Form_Cotizacion2
    Public Property Padre2 As CotizacionUdp
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult

        If DataGridView1.Rows.Count > 0 Then
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

            End If
        Else
            MsgBox("No hay Informacion en el Datagrid, para poder eliminar")
        End If



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim NU, i17, num2, nu2 As Integer


        If TextBox1.Text = "" Then
            MsgBox("FALTA INGRESAR EL ARTICULO")
        Else
            DataGridView1.Rows.Add(1)
            NU = DataGridView1.Rows.Count

            If NU = 1 Then
                DataGridView1.Rows(0).Cells(0).Value = TextBox1.Text

                DataGridView1.Rows(0).Cells(7).Value = NU + 1
                DataGridView1.Rows(0).Cells(1).Value = 1
                DataGridView1.Rows(0).Cells(5).Value = 0
                DataGridView1.Rows(0).Cells(3).Value = "UND"
                DataGridView1.Rows(0).Cells(2).Value = 0
                DataGridView1.Rows(0).Cells(4).Value = 0
                DataGridView1.Rows(0).Cells(6).Value = 0
            Else
                nu2 = DataGridView1.Rows.Count
                For o = 1 To nu2 - 1


                    If DataGridView1.Rows(o).Cells(7).Value < DataGridView1.Rows(o - 1).Cells(7).Value Then
                        num2 = DataGridView1.Rows(o - 1).Cells(7).Value + 1
                        i17 = o

                    End If


                Next


                If i17 = "0" Then
                    MsgBox("SOLO SE PERMITE 5 REGISTROS")
                Else
                    DataGridView1.Rows(i17).Cells(0).Value = TextBox1.Text
                    DataGridView1.Rows(i17).Cells(5).Value = 0
                    DataGridView1.Rows(i17).Cells(7).Value = num2
                    DataGridView1.Rows(i17).Cells(1).Value = 1
                    DataGridView1.Rows(i17).Cells(2).Value = 0
                    DataGridView1.Rows(i17).Cells(4).Value = 0
                    DataGridView1.Rows(i17).Cells(6).Value = 0
                    DataGridView1.Rows(i17).Cells(3).Value = "UND"



                End If


            End If
        End If
        TextBox1.Text = ""

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox2.Text = 1 Then

            Dim va, la As Integer
            Dim sum As Double
            va = DataGridView1.Rows.Count
            Padre2.DataGridView4.Rows.Clear()
            If va > 0 Then
                Padre2.DataGridView4.Rows.Add(va)
                For o = 0 To va - 1
                    Padre2.DataGridView4.Rows(o).Cells(0).Value = DataGridView1.Rows(o).Cells(0).Value
                    Padre2.DataGridView4.Rows(o).Cells(1).Value = DataGridView1.Rows(o).Cells(1).Value
                    Padre2.DataGridView4.Rows(o).Cells(2).Value = DataGridView1.Rows(o).Cells(2).Value
                    Padre2.DataGridView4.Rows(o).Cells(3).Value = DataGridView1.Rows(o).Cells(3).Value
                    Padre2.DataGridView4.Rows(o).Cells(4).Value = DataGridView1.Rows(o).Cells(4).Value
                    Padre2.DataGridView4.Rows(o).Cells(5).Value = DataGridView1.Rows(o).Cells(5).Value
                    Padre2.DataGridView4.Rows(o).Cells(6).Value = DataGridView1.Rows(o).Cells(6).Value
                    Padre2.DataGridView4.Rows(o).Cells(7).Value = DataGridView1.Rows(o).Cells(7).Value
                Next

            End If
            la = DataGridView1.Rows.Count
            For p = 0 To la - 1
                sum = sum + CDbl(DataGridView1.Rows(p).Cells(6).Value)
            Next
            Padre2.TextBox20.Text = Math.Round(sum, 3)
            Close()
        Else
            If TextBox2.Text = 2 Then

                Dim va, va2, la As Integer
                Dim sum As Double
                va = DataGridView1.Rows.Count
                va2 = Padre2.DataGridView5.Rows.Count
                If va2 > 0 Then
                    Padre2.DataGridView5.Rows.Clear()
                End If


                If va > 0 Then
                    Padre2.DataGridView5.Rows.Add(va)
                    For o = 0 To va - 1
                        Padre2.DataGridView5.Rows(o).Cells(0).Value = DataGridView1.Rows(o).Cells(0).Value
                        Padre2.DataGridView5.Rows(o).Cells(1).Value = DataGridView1.Rows(o).Cells(1).Value
                        Padre2.DataGridView5.Rows(o).Cells(2).Value = DataGridView1.Rows(o).Cells(2).Value
                        Padre2.DataGridView5.Rows(o).Cells(3).Value = DataGridView1.Rows(o).Cells(3).Value
                        Padre2.DataGridView5.Rows(o).Cells(4).Value = DataGridView1.Rows(o).Cells(4).Value
                        Padre2.DataGridView5.Rows(o).Cells(5).Value = DataGridView1.Rows(o).Cells(5).Value
                        Padre2.DataGridView5.Rows(o).Cells(6).Value = DataGridView1.Rows(o).Cells(6).Value
                        Padre2.DataGridView5.Rows(o).Cells(7).Value = DataGridView1.Rows(o).Cells(7).Value

                    Next
                End If
                la = DataGridView1.Rows.Count
                For p = 0 To la - 1
                    sum = sum + CDbl(DataGridView1.Rows(p).Cells(6).Value)
                Next
                Padre2.TextBox18.Text = Math.Round(sum, 3)
                Close()
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, fila As Integer
        'Dim cant9 As Double
        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CONSUMO" Then
            DataGridView1.Rows(fila).Cells(4).Value = (Val(DataGridView1.Rows(fila).Cells(2).Value) * (1 + Val(DataGridView1.Rows(fila).Cells(1).Value) / 100))
            Me.DataGridView1.Columns(4).DefaultCellStyle.Format = "n3"
            DataGridView1.Rows(fila).Cells(6).Value = Val(DataGridView1.Rows(fila).Cells(4).Value) * (Val(DataGridView1.Rows(fila).Cells(5).Value))
            Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n3"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "% MERMA" Then
            DataGridView1.Rows(fila).Cells(4).Value = (Val(DataGridView1.Rows(fila).Cells(2).Value) * (1 + Val(DataGridView1.Rows(fila).Cells(1).Value) / 100))
            Me.DataGridView1.Columns(4).DefaultCellStyle.Format = "n3"
            DataGridView1.Rows(fila).Cells(6).Value = Val(DataGridView1.Rows(fila).Cells(4).Value) * (Val(DataGridView1.Rows(fila).Cells(5).Value))
            Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n3"

        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "C.UNITARIO" Then
            DataGridView1.Rows(fila).Cells(6).Value = Val(DataGridView1.Rows(fila).Cells(4).Value) * (Val(DataGridView1.Rows(fila).Cells(5).Value))
            Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n3"

        End If
    End Sub

    Private Sub Form_Cotizacion2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class