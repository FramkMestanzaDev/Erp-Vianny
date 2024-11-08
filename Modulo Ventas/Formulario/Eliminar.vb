Public Class Eliminar
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        Dim ui As Integer
        ui = Nota_Pedido.DataGridView1.Rows.Count
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Dim nu As Integer
            nu = TextBox1.Text
            If nu > ui Then
                MsgBox("EL ITEMS A ELIMINAR NO EXISTE")
            Else
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(0).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(1).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(2).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(3).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(4).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(5).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(6).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(7).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(8).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(9).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(10).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(11).Value = ""
                Nota_Pedido.DataGridView1.Rows(nu - 1).Cells(12).Value = ""
                Nota_Pedido.DataGridView1.Rows.RemoveAt(nu - 1)

                Dim j As New Integer
                j = Nota_Pedido.DataGridView1.Rows.Count - 1
                If j > 0 Then
                    For i = 0 To nu - 1
                        Nota_Pedido.DataGridView1.Rows(i).Cells(0).Value = i
                    Next
                End If

                Dim i1, a9 As Integer
                Dim cant9, sum9, cant8, sum8, cant7, sum7 As Double
                i1 = Nota_Pedido.DataGridView1.Rows.Count

                For a9 = 0 To i1 - 1
                    cant9 = Val(Nota_Pedido.DataGridView1.Rows(a9).Cells(10).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(Nota_Pedido.DataGridView1.Rows(a9).Cells(11).Value)
                    sum8 = cant8 + Val(sum8)
                    cant7 = Val(Nota_Pedido.DataGridView1.Rows(a9).Cells(12).Value)
                    sum7 = cant7 + Val(sum7)

                Next
                Nota_Pedido.TextBox15.Text = sum9.ToString("N2")
                Nota_Pedido.TextBox16.Text = sum8.ToString("N2")
                Nota_Pedido.TextBox17.Text = sum7.ToString("N2")
                Me.Close()

            End If

        End If
    End Sub
End Class