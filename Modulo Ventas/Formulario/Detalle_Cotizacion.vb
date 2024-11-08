Public Class Detalle_Cotizacion
    Dim DT As New DataTable
    Private Sub Detalle_Cotizacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If TextBox1.Text = 1 Then
            Dim VG As New vTela
            Dim VA As New ftela

            VG.gid_cotizacion = TextBox2.Text
            DT = VA.buscar_tela_codigo(VG)
            If DT.Rows.Count <> 0 Then
                DataGridView1.DataSource = DT
                DataGridView1.Columns(0).Width = 120
                DataGridView1.Columns(1).Width = 420
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(0).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(1).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(2).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(3).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(4).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                DataGridView1.Columns(5).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
            End If
        Else
            If TextBox1.Text = 2 Then
                Dim VG As New vAvios_Acabados
                Dim VA As New favios_acabados

                VG.gid_cotizacion = TextBox2.Text
                DT = VA.buscar_Avios_Acabados_codigo(VG)
                If DT.Rows.Count <> 0 Then
                    DataGridView1.DataSource = DT
                    DataGridView1.Columns(0).Width = 120
                    DataGridView1.Columns(1).Width = 420
                    DataGridView1.Columns(0).Frozen = True
                    DataGridView1.Columns(1).Frozen = True
                    DataGridView1.Columns(0).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                    DataGridView1.Columns(1).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                    DataGridView1.Columns(2).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                    DataGridView1.Columns(3).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                    DataGridView1.Columns(4).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                    DataGridView1.Columns(5).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                End If
            Else
                If TextBox1.Text = 3 Then
                    Dim VG As New vAvios_Costura
                    Dim VA As New favios_costura

                    VG.gid_cotizacion = TextBox2.Text
                    DT = VA.buscar_Avios_Costura_codigo(VG)
                    If DT.Rows.Count <> 0 Then
                        DataGridView1.DataSource = DT
                        DataGridView1.Columns(0).Width = 120
                        DataGridView1.Columns(1).Width = 420
                        DataGridView1.Columns(0).Frozen = True
                        DataGridView1.Columns(1).Frozen = True
                        DataGridView1.Columns(0).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                        DataGridView1.Columns(1).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                        DataGridView1.Columns(2).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                        DataGridView1.Columns(3).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                        DataGridView1.Columns(4).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                        DataGridView1.Columns(5).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                    End If
                Else
                    If TextBox1.Text = 4 Then
                        Dim VG As New vLavados
                        Dim VA As New flavados

                        VG.gid_cotizacion = TextBox2.Text
                        DT = VA.buscar_Lavados_codigo(VG)
                        If DT.Rows.Count <> 0 Then
                            DataGridView1.DataSource = DT
                            DataGridView1.Columns(0).Width = 370
                            DataGridView1.Columns(0).Frozen = True
                            DataGridView1.Columns(1).Frozen = True
                            DataGridView1.Columns(0).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                            DataGridView1.Columns(1).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                            DataGridView1.Columns(2).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                            DataGridView1.Columns(3).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                            DataGridView1.Columns(4).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                        End If
                    Else
                        If TextBox1.Text = 5 Then
                            Dim VG As New Aplicaciones
                            Dim VA As New faplicaciones

                            VG.gid_cotizacion = TextBox2.Text
                            DT = VA.buscar_Aplicaciones_codigo(VG)
                            If DT.Rows.Count <> 0 Then
                                DataGridView1.DataSource = DT
                                DataGridView1.Columns(0).Width = 370

                                DataGridView1.Columns(0).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                                DataGridView1.Columns(1).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                                DataGridView1.Columns(2).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                                DataGridView1.Columns(3).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                                DataGridView1.Columns(4).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                            End If
                        Else
                            Dim VG As New vManoObra
                            Dim VA As New fmano_obra

                            VG.gid_cotizacion = TextBox2.Text
                            DT = VA.buscar_ManoObra_codigo(VG)
                            If DT.Rows.Count <> 0 Then
                                DataGridView1.DataSource = DT
                                DataGridView1.Columns(0).Width = 370

                                DataGridView1.Columns(0).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                                DataGridView1.Columns(1).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                                DataGridView1.Columns(2).DefaultCellStyle.Alignment = ContentAlignment.TopCenter
                                DataGridView1.Columns(3).DefaultCellStyle.Alignment = ContentAlignment.TopCenter

                            End If
                        End If


                    End If

                    End If

            End If

        End If
    End Sub
End Class