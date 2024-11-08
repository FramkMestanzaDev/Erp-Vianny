Public Class Productos_Hilo
    Dim gk As DataTable

    Private Sub Productos_Hilo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TG As New FSTOCK_PAR
        Dim HJ As New VSTOCK_PAR
        Dim K As Integer


        Try


            'Label5.Text = Nsalida.Label10.Text
            HJ.gCCIA = Label6.Text
            HJ.gALMACEN = Label3.Text

            HJ.gMODAL = Label7.Text

            HJ.gano = Label5.Text
            gk = TG.CaeSoft_ReporteStockFisico_hilo(HJ)
            DataGridView1.DataSource = gk

            DataGridView1.Columns(1).Width = 142
            DataGridView1.Columns(2).Width = 330
            DataGridView1.Columns(3).Width = 80
            DataGridView1.Columns(4).Width = 70
            DataGridView1.Columns(5).Width = 100
            DataGridView1.Columns(6).Width = 70
            K = DataGridView1.Rows.Count

        Catch ex As Exception
            MsgBox("NO EXISTE ESTOCK EN ESTE PERIODO")
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Label4.Text = 200 Then
            Dim i, I2, SUMA As Integer
            Dim suma2 As String
            suma2 = ""
            i = DataGridView1.Rows.Count
            For b = 0 To i - 1
                If Me.DataGridView1.Rows(b).Cells(0).Value = True Then
                    NiaHilo.DataGridView1.Rows.Add(1)
                    I2 = NiaHilo.DataGridView1.Rows.Count
                    If I2 = 1 Then
                        NiaHilo.DataGridView1.Rows(0).Cells(0).Value = "001"
                        NiaHilo.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                        NiaHilo.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                        Select Case Label3.Text
                            Case "41" : NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "KGS"
                            Case "03" : NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "CAJ"
                            Case "42" : NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "BOL"
                            Case "43" : NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                            Case "44" : NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                            Case "13" : NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                            Case "14" : NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                            Case "11" : NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                            Case "05" : NiaHilo.DataGridView1.Rows(0).Cells(5).Value = "UND"

                        End Select
                        NiaHilo.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(b).Cells(5).Value

                        NiaHilo.DataGridView1.Rows(0).Cells(8).Value = DataGridView1.Rows(b).Cells(3).Value
                        NiaHilo.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(b).Cells(6).Value
                        NiaHilo.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                        'DataGridView1.Rows(b).ReadOnly = True
                        'Me.DataGridView1.Rows(b).Cells(0).Value = True
                        'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.BurlyWood
                    Else
                        If I2 > 1 Then

                            SUMA = NiaHilo.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                            'If I2 > 998 Then
                            '    MsgBox(NiaHilo.DataGridView1.Rows(I2 - 2).Cells(0).Value)
                            '    MsgBox(SUMA)
                            '    MsgBox(b)
                            'End If
                            Select Case SUMA.ToString.Length
                                Case "1" : suma2 = "000" & "" & SUMA
                                Case "2" : suma2 = "00" & "" & SUMA
                                Case "3" : suma2 = "0" & "" & SUMA
                                Case "4" : suma2 = SUMA

                            End Select
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(0).Value = suma2
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                            Select Case Label3.Text
                                Case "41" : NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "KGS"
                                Case "03" : NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "CAJ"
                                Case "42" : NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "BOL"
                                Case "43" : NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                                Case "44" : NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                                Case "13" : NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                                Case "14" : NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                                Case "11" : NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                                Case "05" : NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                            End Select

                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(b).Cells(5).Value
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(8).Value = DataGridView1.Rows(b).Cells(3).Value
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(7).Value = DataGridView1.Rows(b).Cells(6).Value
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                        End If
                    End If
                    'DataGridView1.Rows(b).ReadOnly = True
                    'Me.DataGridView1.Rows(b).Cells(0).Value = True
                    'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.Red

                End If

            Next
            NiaHilo.DataGridView1.Columns(0).Frozen = True
            NiaHilo.DataGridView1.Columns(1).Frozen = True
            NiaHilo.DataGridView1.Columns(2).Frozen = True


        End If
        If Label4.Text = 2 Then
            Dim i, I2, SUMA As Integer
            Dim suma2 As String
            suma2 = ""
            i = DataGridView1.Rows.Count
            For b = 0 To i - 1
                If Me.DataGridView1.Rows(b).Cells(0).Value = True Then
                    NsaHilo.DataGridView1.Rows.Add(1)
                    I2 = NsaHilo.DataGridView1.Rows.Count
                    If I2 = 1 Then
                        NsaHilo.DataGridView1.Rows(0).Cells(0).Value = "001"
                        NsaHilo.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                        NsaHilo.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                        Select Case Label3.Text
                            Case "41" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "KGS"
                            Case "03" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "CON"
                            Case "42" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "CON"
                            Case "43" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                            Case "44" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                        End Select
                        NsaHilo.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(b).Cells(5).Value
                        NsaHilo.DataGridView1.Rows(0).Cells(10).Value = DataGridView1.Rows(b).Cells(5).Value
                        NsaHilo.DataGridView1.Rows(0).Cells(8).Value = DataGridView1.Rows(b).Cells(3).Value
                        NsaHilo.DataGridView1.Rows(0).Cells(11).Value = DataGridView1.Rows(b).Cells(6).Value
                        NsaHilo.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                        'DataGridView1.Rows(b).ReadOnly = True
                        'Me.DataGridView1.Rows(b).Cells(0).Value = True
                        'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.BurlyWood
                    Else
                        If I2 > 1 Then
                            SUMA = NsaHilo.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                            Select Case SUMA.ToString.Length
                                Case "1" : suma2 = "00" & "" & SUMA
                                Case "2" : suma2 = "0" & "" & SUMA
                                Case "3" : suma2 = SUMA
                            End Select
                            NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(0).Value = suma2
                            NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                            NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                            Select Case Label3.Text
                                Case "41" : NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "KGS"
                                Case "03" : NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "CON"
                                Case "42" : NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "CON"
                                Case "43" : NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                                Case "44" : NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                            End Select

                            NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(b).Cells(5).Value
                            NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(10).Value = DataGridView1.Rows(b).Cells(5).Value
                            NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(8).Value = DataGridView1.Rows(b).Cells(3).Value
                            NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(11).Value = DataGridView1.Rows(b).Cells(6).Value
                            NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                        End If
                    End If
                    'DataGridView1.Rows(b).ReadOnly = True
                    'Me.DataGridView1.Rows(b).Cells(0).Value = True
                    'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.Red

                End If

            Next
            NsaHilo.DataGridView1.Columns(0).Frozen = True
            NsaHilo.DataGridView1.Columns(1).Frozen = True
            NsaHilo.DataGridView1.Columns(2).Frozen = True


        End If
        If Label4.Text = 1 Then
            Dim i, I2, SUMA As Integer
            Dim suma2 As String
            suma2 = ""
            i = DataGridView1.Rows.Count
            For b = 0 To i - 1
                If Me.DataGridView1.Rows(b).Cells(0).Value = True Then
                    Guia_hilo.DataGridView1.Rows.Add(1)
                    I2 = Guia_hilo.DataGridView1.Rows.Count
                    If I2 = 1 Then
                        Guia_hilo.DataGridView1.Rows(0).Cells(0).Value = "001"
                        Guia_hilo.DataGridView1.Rows(0).Cells(21).Value = "001"
                        Guia_hilo.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                        Guia_hilo.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                        Select Case Label3.Text
                            Case "41" : Guia_hilo.DataGridView1.Rows(0).Cells(5).Value = "KGS"
                            Case "03" : Guia_hilo.DataGridView1.Rows(0).Cells(5).Value = "CON"
                            Case "42" : Guia_hilo.DataGridView1.Rows(0).Cells(5).Value = "CON"
                            Case "43" : Guia_hilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                            Case "44" : Guia_hilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                        End Select
                        Guia_hilo.DataGridView1.Rows(0).Cells(13).Value = 0
                        Guia_hilo.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(b).Cells(3).Value
                        'Guia_hilo.DataGridView1.Rows(0).Cells(10).Value = DataGridView1.Rows(b).Cells(3).Value
                        Guia_hilo.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(b).Cells(5).Value
                        Guia_hilo.DataGridView1.Rows(0).Cells(8).Value = DataGridView1.Rows(b).Cells(6).Value
                        Guia_hilo.DataGridView1.Rows(0).Cells(11).Value = DataGridView1.Rows(b).Cells(5).Value
                        Guia_hilo.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                        'DataGridView1.Rows(b).ReadOnly = True
                        'Me.DataGridView1.Rows(b).Cells(0).Value = True
                        'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.BurlyWood
                    Else
                        If I2 > 1 Then

                            SUMA = Guia_hilo.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1

                            Select Case SUMA.ToString.Length
                                Case "1" : suma2 = "00" & "" & SUMA
                                Case "2" : suma2 = "0" & "" & SUMA
                                Case "3" : suma2 = SUMA
                            End Select
                            Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(13).Value = 0
                            Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(0).Value = suma2
                            Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(21).Value = suma2
                            Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                            Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                            Select Case Label3.Text
                                Case "41" : Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "KGS"
                                Case "03" : Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "CON"
                                Case "42" : Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "CON"
                                Case "43" : Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                                Case "44" : Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                            End Select
                            Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(8).Value = DataGridView1.Rows(b).Cells(6).Value
                            'Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(10).Value = DataGridView1.Rows(b).Cells(3).Value
                            Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(b).Cells(3).Value
                            Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(7).Value = DataGridView1.Rows(b).Cells(5).Value
                            Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(11).Value = DataGridView1.Rows(b).Cells(5).Value
                            Guia_hilo.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                        End If
                    End If
                    'DataGridView1.Rows(b).ReadOnly = True
                    'Me.DataGridView1.Rows(b).Cells(0).Value = True
                    'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.Red

                End If

            Next
            Guia_hilo.DataGridView1.Columns(0).Frozen = True
            Guia_hilo.DataGridView1.Columns(1).Frozen = True
            Guia_hilo.DataGridView1.Columns(2).Frozen = True


        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim J As Integer

        J = DataGridView1.Rows.Count
        If CheckBox1.Checked = True Then
            For i = 0 To J - 1
                DataGridView1.Rows(i).Cells(0).Value = True
            Next
        Else
            For i = 0 To J - 1
                DataGridView1.Rows(i).Cells(0).Value = False
            Next
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar_PARTIDA()
    End Sub
    Private Sub buscar_PARTIDA()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(gk.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "LOTE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 142
                DataGridView1.Columns(2).Width = 330
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(5).Width = 100
                DataGridView1.Columns(6).Width = 70
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar_Producto()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(gk.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim textoBusqueda As String = TextBox2.Text
            dv.RowFilter = $"PRODUCTO LIKE '%{textoBusqueda}%' "
            If dv.Count <> 0 Then




                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 142
                DataGridView1.Columns(2).Width = 330
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(5).Width = 100
                DataGridView1.Columns(6).Width = 70
            Else

                DataGridView1.DataSource = Nothing
            End If



        Catch ex As Exception


        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar_Producto()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) 
        Dim ex As New Exportar
        ex.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class