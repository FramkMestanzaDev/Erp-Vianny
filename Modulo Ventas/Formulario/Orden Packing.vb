Public Class Orden_Packing
    Dim dt As New DataTable
    Sub ACTUALIZAR()

        Dim hu As New fpedidopacking
        Dim hu1 As New vpackign

        hu1.galmacen = Label3.Text
        hu1.gestado = 1
        hu1.gpartida = Label4.Text
        hu1.gCalidad = Label10.Text
        dt = hu.buscar_ped_packin_estado(hu1)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(12).Visible = False
            TextBox3.Text = DataGridView1.Rows(0).Cells(4).Value
            TextBox2.Text = DataGridView1.Rows(0).Cells(3).Value
            DataGridView1.Columns(1).Width = 90
            DataGridView1.Columns(2).Width = 180
            DataGridView1.Columns(3).Width = 300
            DataGridView1.Columns(4).Width = 90
        Else
            MsgBox("NO HAY ORDEN DE PACKIN PENDIENTES POR ADJUNTAR")
        End If
    End Sub
    Private Sub Orden_Packing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = ""
        dt.Clear()
        ACTUALIZAR()
        Dim l As Integer
        l = DataGridView1.Rows.Count
        For i = 0 To l - 1
            If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" And Trim(DataGridView1.Rows(i).Cells(13).Value) = Trim(TextBox6.Text) Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Rows(i).Cells(0).Value = True
                DataGridView1.Rows(i).ReadOnly = True
            Else
                If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" And Trim(DataGridView1.Rows(i).Cells(13).Value) <> Trim(TextBox6.Text) Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Black
                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    DataGridView1.Rows(i).ReadOnly = True
                End If
            End If
            Dim sumatotal As Double
            sumatotal = 0

            For Each dgvrow As DataGridViewRow In DataGridView1.Rows
                If dgvrow.Cells("s").Value = True Then
                    sumatotal = sumatotal + CDbl(dgvrow.Cells("PESO").Value)
                End If
            Next
            TextBox5.Text = sumatotal
        Next
    End Sub
    Dim DT5 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a, i2, i3 As Integer
        Dim valor As Double
        Dim rollo As String
        Dim jlk As New fpedidopacking
        Dim kll As New vpackign
        i = DataGridView1.Rows.Count

        'For a = 0 To i - 1
        '    If Me.DataGridView1.Rows(a).Cells(0).Value = True Then

        '        valor = valor + DataGridView1.Rows(a).Cells(2).Value

        '    End If

        'Next
        If Convert.ToDouble(TextBox5.Text) <= Convert.ToDouble(TextBox4.Text) Then
            For a = 0 To DataGridView1.Rows.Count - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                    Nota_Pedido.DataGridView7.Rows.Add()
                    i3 = Nota_Pedido.DataGridView7.Rows.Count
                    Nota_Pedido.DataGridView7.Rows(i3 - 1).Cells(0).Value = DataGridView1.Rows(a).Cells(7).Value
                    Nota_Pedido.DataGridView7.Rows(i3 - 1).Cells(1).Value = DataGridView1.Rows(a).Cells(9).Value
                    Nota_Pedido.DataGridView7.Rows(i3 - 1).Cells(2).Value = DataGridView1.Rows(a).Cells(2).Value
                    Nota_Pedido.DataGridView7.Rows(i3 - 1).Cells(3).Value = DataGridView1.Rows(a).Cells(5).Value
                    If a = 0 Then
                        MsgBox(a)
                        rollo = Trim(DataGridView1.Rows(a).Cells(7).Value)
                    Else
                        rollo = rollo + "," + Trim(DataGridView1.Rows(a).Cells(7).Value)
                    End If

                    kll.gpartida = Trim(TextBox1.Text)
                    kll.gID = DataGridView1.Rows(a).Cells(7).Value
                        kll.gselec = 1
                        kll.galmacen = Label3.Text
                        kll.gestado = 1
                        kll.gCalidad = Label10.Text
                        kll.gNpedido = Trim(TextBox6.Text)
                        jlk.actualizar_rollos(kll)
                        '    'Guia_remision.DataGridView5.Rows.Add(1)
                        '    'i2 = Guia_remision.DataGridView5.Rows.Count

                        '    If i2 = 1 Then
                        '        Guia_remision.DataGridView5.Rows(0).Cells(0).Value = DataGridView1.Rows(a).Cells(7).Value
                        '        Guia_remision.DataGridView5.Rows(0).Cells(1).Value = DataGridView1.Rows(a).Cells(9).Value
                        '        Guia_remision.DataGridView5.Rows(0).Cells(2).Value = DataGridView1.Rows(a).Cells(2).Value
                        '        Guia_remision.DataGridView5.Rows(0).Cells(3).Value = DataGridView1.Rows(a).Cells(5).Value
                        '    Else
                        '        If i2 > 1 Then
                        '            Guia_remision.DataGridView5.Rows(i2 - 1).Cells(0).Value = DataGridView1.Rows(a).Cells(7).Value
                        '            Guia_remision.DataGridView5.Rows(i2 - 1).Cells(1).Value = DataGridView1.Rows(a).Cells(9).Value
                        '            Guia_remision.DataGridView5.Rows(i2 - 1).Cells(2).Value = DataGridView1.Rows(a).Cells(2).Value
                        '            Guia_remision.DataGridView5.Rows(i2 - 1).Cells(3).Value = DataGridView1.Rows(a).Cells(5).Value
                        '        End If
                        '    End If
                        i3 = i3 + 1
                    End If
            Next
            Nota_Pedido.DataGridView1.Rows(Label9.Text).Cells(5).Value = rollo
            Nota_Pedido.DataGridView1.Rows(Label9.Text).Cells(8).Value = TextBox5.Text
            Close()
        Else
            MsgBox("LA CANTIDAD VALIDADA EN EL PACKING NO COINCIDE CON LA CANTIDAD SOLICITADA EN LA GUIA")
        End If



    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim i, a, i2 As Integer
        Dim jlk As New fpedidopacking
        Dim kll As New vpackign
        i = DataGridView1.Rows.Count

        If i > 0 Then
            For a = 0 To i - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                    kll.gpartida = Trim(TextBox1.Text)
                    kll.gID = DataGridView1.Rows(a).Cells(7).Value
                    kll.gselec = 0
                    kll.gestado = 1
                    kll.galmacen = Label3.Text
                    kll.gCalidad = Label10.Text
                    kll.gNpedido = ""
                    jlk.actualizar_rollos(kll)
                    Dim k As Integer

                    k = Guia_remision.DataGridView5.Rows.Count
                    For a1 = 0 To k - 1
                        If DataGridView1.Rows(a).Cells(7).Value = Guia_remision.DataGridView5.Rows(a1).Cells(0).Value Then
                            Guia_remision.DataGridView5.Rows.RemoveAt(a1)
                        End If

                    Next


                End If
            Next
            ACTUALIZAR()
        Else
            MsgBox("NO HAY ELEMENTOS PARA ELIMINAR")
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick




    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick


    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Dim j As Integer
            j = DataGridView1.Rows.Count
            For i = 0 To j - 1
                DataGridView1.Rows(i).Cells(0).Value = True
            Next

        End If
    End Sub



    Private Sub DataGridView1_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseUp
        If e.ColumnIndex = 0 Then
            '    If DataGridView1.Rows(e.RowIndex).Cells("s").Value = False Then
            '        DataGridView1.Rows(e.RowIndex).Cells("s").Value = True
            '        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White
            '        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Red
            '    Else
            '        DataGridView1.Rows(e.RowIndex).Cells("s").Value = False
            '        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Black
            '        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
            '    End If
            Dim sumatotal As Double
            sumatotal = 0

            For Each dgvrow As DataGridViewRow In DataGridView1.Rows
                If dgvrow.Cells("s").Value = True Then
                    sumatotal = sumatotal + CDbl(dgvrow.Cells("PESO").Value)
                End If
            Next
            TextBox5.Text = sumatotal
        End If
    End Sub

    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        'Dim l As Integer
        'l = DataGridView1.Rows.Count
        'If l > 0 Then
        '    If e.ColumnIndex = 0 And Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value) = "0" Then
        '        If DataGridView1.Rows(e.RowIndex).Cells(0).Value = True Then
        '            'DataGridView1.Rows(e.RowIndex).Cells(0).Value = True
        '            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White
        '            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Red
        '        Else
        '            'DataGridView1.Rows(e.RowIndex).Cells(0).Value = False
        '            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Black
        '            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
        '        End If
        '        Dim sumatotal As Double
        '        sumatotal = 0

        '        For Each dgvrow As DataGridViewRow In DataGridView1.Rows
        '            If dgvrow.Cells(0).Value = True Then
        '                sumatotal = sumatotal + CDbl(dgvrow.Cells(9).Value)
        '            End If
        '        Next
        '        TextBox5.Text = sumatotal
        '    End If
        'End If

    End Sub

    Private Sub Orden_Packing_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

    End Sub

    Private Sub Orden_Packing_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Dim respuesta As DialogResult
        'respuesta = MessageBox.Show("Al Salir de esta manera de quitara las selecciones que a realizado, esta Seguro?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        'If (respuesta = Windows.Forms.DialogResult.Yes) Then
        '    Dim i, a, i2 As Integer
        '    Dim jlk As New fpedidopacking
        '    Dim kll As New vpackign
        '    i = DataGridView1.Rows.Count

        '    If i > 0 Then
        '        For a = 0 To i - 1
        '            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
        '                kll.gpartida = Trim(TextBox1.Text)
        '                kll.gID = DataGridView1.Rows(a).Cells(7).Value
        '                kll.gselec = 0
        '                kll.gestado = 1
        '                kll.galmacen = Label3.Text
        '                kll.gCalidad = Label10.Text
        '                kll.gNpedido = ""
        '                jlk.actualizar_rollos(kll)
        '                Dim k As Integer
        '            End If
        '        Next
        '    End If
        'End If
    End Sub
End Class