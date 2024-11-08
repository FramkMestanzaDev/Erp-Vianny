Public Class Form2
    Dim dt As New DataTable
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim g As New FCOTIZACION
        Dim g1 As New VCOTIZACION
        Dim UA As Integer
        If TextBox3.Text = 1 Then
            g1.gvendedor = Op_Manufactura.TextBox10.Text
        Else
            If TextBox3.Text = 2 Then
                g1.gvendedor = Orden_Produccion.TextBox11.Text
            Else
                If TextBox3.Text = 3 Then
                    g1.gvendedor = Nota_Pedido.TextBox9.Text
                End If
            End If


        End If



            dt = g.buscar_vendedor_cotizacion(g1)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt

            DataGridView1.Columns(1).Width = 200
            DataGridView1.Columns(3).Width = 230
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(9).Visible = False
            UA = DataGridView1.Rows.Count
            'For i = 0 To UA - 1
            '    If DataGridView1.Rows(i).Cells(8).Value = 1 Then
            '        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkSlateBlue
            '    End If
            'Next


        Else
            MsgBox("NO HY INFORMACION PARA MOSTRAR")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a, conta As Integer
        i = DataGridView1.Rows.Count
        a = 0
        conta = 0



        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                conta = conta + 1
            End If
        Next
        If conta > 1 Then
            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
        Else

            For a = 0 To i - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                    If DataGridView1.Rows(a).Cells(8).Value = 0 Then
                        MsgBox("LA COTIZACION NO ESTA APROBADO")
                    Else
                        If DataGridView1.Rows(a).Cells(8).Value = 1 Then
                            If TextBox3.Text = 1 Then
                                Op_Manufactura.TextBox7.Text = DataGridView1.Rows(a).Cells(7).Value
                            Else
                                If TextBox3.Text = 2 Then
                                    'Orden_Produccion.TextBox18.Text = DataGridView1.Rows(a).Cells(7).Value
                                Else

                                End If

                            End If

                        End If
                    End If
                End If
            Next

            Close()
        End If
    End Sub
End Class