Public Class Lista_ARti
    Dim DT As DataTable
    Private Sub Lista_ARti_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim GH As New fcliente
        Dim gg As New vcliente
        gg.galmacen = Label2.Text
        gg.gccia = Label3.Text
        DT = GH.buscar_codigo_almacen2(gg)
        DataGridView1.DataSource = DT
        DataGridView1.Columns(1).Width = 150
        DataGridView1.Columns(2).Width = 350

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 150
                DataGridView1.Columns(2).Width = 350

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Label2.Text = "06" Then
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

                        NotaIngreso.TextBox20.Text = Trim(DataGridView1.Rows(a).Cells(1).Value)

                    End If
                Next

            End If
            NotaIngreso.TextBox20.Select()
        Else
            Dim i, I2, SUMA As Integer

            i = DataGridView1.Rows.Count
            For a = 0 To i - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then

                    NiaHilo.DataGridView1.Rows.Add(1)
                    I2 = NiaHilo.DataGridView1.Rows.Count

                    If I2 = 1 Then


                        NiaHilo.DataGridView1.Rows(0).Cells(0).Value = 1

                        NiaHilo.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(a).Cells(1).Value
                        NiaHilo.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(a).Cells(2).Value
                        NiaHilo.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(a).Cells(3).Value
                        If Trim(DataGridView1.Rows(a).Cells(1).Value) = "TELFRENC00405000595" Or Trim(DataGridView1.Rows(a).Cells(1).Value) = "TELFRENC00023000591" Then
                            NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "RLL"
                        Else
                            NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "CON"
                        End If




                    Else
                        If I2 > 1 Then

                            SUMA = NiaHilo.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(0).Value = SUMA

                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(a).Cells(1).Value
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(a).Cells(2).Value
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(7).Value = DataGridView1.Rows(a).Cells(3).Value

                            If Trim(DataGridView1.Rows(a).Cells(1).Value) = "TELFRENC00405000595" Or Trim(DataGridView1.Rows(a).Cells(1).Value) = "TELFRENC00023000591" Then
                                NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "RLL"
                            Else
                                NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "CON"
                            End If

                        End If
                    End If
                    'DataGridView1.Rows(a).ReadOnly = True
                    'Me.DataGridView1.Rows(a).Cells(0).Value = False
                    'DataGridView1.Rows(a).DefaultCellStyle.BackColor = Color.Red
                End If
            Next
        End If


        Close()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class