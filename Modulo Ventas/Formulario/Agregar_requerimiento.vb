Imports System.Data.SqlClient
Class Agregar_requerimiento
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Agregar_requerimiento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("select id_requeminieto AS REQUERIMIENTO,case when area= 1 then 'CONFECCION' ELSE 'ACABADOS' END AS AREA,ruc AS RUC,rsocial AS TALLER,op AS OP, PO AS PO,cliente AS CLIENTE,producto AS PRODUCTO,corte AS CORTE, fdespacho AS DESPACHO from requerimiento_avios_cabecera   WHERE   estado ='3'", conx)
        Dim da As New DataTable
        cmd.Fill(da)
        DataGridView1.DataSource = da
        Dim ol As Integer

        ol = DataGridView1.Rows.Count
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 85
        DataGridView1.Columns(5).Width = 83
        DataGridView1.Columns(6).Width = 83
        DataGridView1.Columns(9).Width = 75
        DataGridView1.Columns(1).ReadOnly = True
        DataGridView1.Columns(2).ReadOnly = True
        DataGridView1.Columns(3).ReadOnly = True
        DataGridView1.Columns(4).ReadOnly = True
        DataGridView1.Columns(5).ReadOnly = True
        DataGridView1.Columns(6).ReadOnly = True
        DataGridView1.Columns(7).ReadOnly = True
        DataGridView1.Columns(8).ReadOnly = True
        DataGridView1.Columns(9).ReadOnly = True
        DataGridView1.Columns(10).ReadOnly = True

    End Sub



    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.ColumnIndex = 1 Then
            detalle_requerimiento.Label1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            detalle_requerimiento.Label2.Text = e.RowIndex
            detalle_requerimiento.Show()
        End If
    End Sub
    Dim Rsr3, Rsr21 As SqlDataReader

    Private Sub Agregar_requerimiento_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim p, p1 As Integer
        p = DataGridView1.Rows.Count
        If Label6.Text = 1 Then
            For i = 0 To p - 1
                If DataGridView1.Rows(i).Cells(0).Value = True Then
                    abrir()
                    Dim i51, i52 As Integer
                    i51 = 0
                    i52 = NsaHilo.DataGridView1.Rows.Count
                    Dim sql102213 As String = "select * from requerimiento_avios_detalle where id_cabecera ='" + Trim(DataGridView1.Rows(i).Cells(1).Value) + "'"
                    Dim cmd102213 As New SqlCommand(sql102213, conx)
                    Rsr3 = cmd102213.ExecuteReader()
                    While Rsr3.Read()
                        NsaHilo.DataGridView1.Rows.Add()
                        NsaHilo.DataGridView1.Rows(i51 + i52).Cells(0).Value = i51 + 1 + i52
                        NsaHilo.DataGridView1.Rows(i51 + i52).Cells(2).Value = Rsr3(3)
                        NsaHilo.DataGridView1.Rows(i51 + i52).Cells(3).Value = Rsr3(4)
                        NsaHilo.DataGridView1.Rows(i51 + i52).Cells(4).Value = Rsr3(8)
                        NsaHilo.DataGridView1.Rows(i51 + i52).Cells(5).Value = Rsr3(6)
                        NsaHilo.DataGridView1.Rows(i51 + i52).Cells(6).Value = Rsr3(8)
                        NsaHilo.DataGridView1.Rows(i51 + i52).Cells(11).Value = Rsr3(6)
                        NsaHilo.DataGridView1.Rows(i51 + i52).Cells(14).Value = Rsr3(0)
                        NsaHilo.DataGridView1.Rows(i51 + i52).Cells(15).Value = Rsr3(1)


                        i51 = i51 + 1
                    End While
                    Rsr3.Close()
                End If





            Next

            p1 = Guia_remision.DataGridView1.Rows.Count

            For k = 0 To p1 - 1
                Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_ReporteStockFisico '" + Trim(Label3.Text) + "','" + Trim(Label4.Text) + "','" + Trim(Label5.Text) + "','" + Trim(Label4.Text) + "0101','" + Trim(Label4.Text) + "1231','" + NsaHilo.DataGridView1.Rows(k).Cells(2).Value + "','" + NsaHilo.DataGridView1.Rows(k).Cells(2).Value + "',NULL, 1"
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr21 = cmd1021.ExecuteReader()

                If Rsr21.Read() Then
                    Guia_remision.DataGridView1.Rows(k).Cells(10).Value = Rsr21(10)
                Else
                    Guia_remision.DataGridView1.Rows(k).Cells(10).Value = 0
                End If

                Rsr21.Close()
            Next

        Else
            If Label6.Text = 2 Then
                For i = 0 To p - 1
                    If DataGridView1.Rows(i).Cells(0).Value = True Then
                        abrir()
                        Dim i51, i52 As Integer
                        i51 = 0
                        i52 = Guia_remision.DataGridView1.Rows.Count
                        Dim sql102213 As String = "select * from requerimiento_avios_detalle where id_cabecera ='" + Trim(DataGridView1.Rows(i).Cells(1).Value) + "'"
                        Dim cmd102213 As New SqlCommand(sql102213, conx)
                        Rsr3 = cmd102213.ExecuteReader()
                        While Rsr3.Read()
                            Guia_remision.DataGridView1.Rows.Add()
                            Guia_remision.DataGridView1.Rows(i51 + i52).Cells(0).Value = i51 + 1 + i52
                            Guia_remision.DataGridView1.Rows(i51 + i52).Cells(2).Value = Rsr3(3)
                            Guia_remision.DataGridView1.Rows(i51 + i52).Cells(3).Value = Rsr3(4)
                            Guia_remision.DataGridView1.Rows(i51 + i52).Cells(4).Value = Rsr3(8)
                            Guia_remision.DataGridView1.Rows(i51 + i52).Cells(5).Value = Rsr3(6)
                            Guia_remision.DataGridView1.Rows(i51 + i52).Cells(7).Value = Rsr3(8)
                            Guia_remision.DataGridView1.Rows(i51 + i52).Cells(8).Value = Rsr3(6)
                            Guia_remision.DataGridView1.Rows(i51 + i52).Cells(14).Value = Rsr3(0)
                            Guia_remision.DataGridView1.Rows(i51 + i52).Cells(15).Value = Rsr3(1)


                            i51 = i51 + 1
                        End While
                        Rsr3.Close()
                    End If





                Next

                p1 = NsaHilo.DataGridView1.Rows.Count

                For k = 0 To p1 - 1
                    Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_ReporteStockFisico '" + Trim(Label3.Text) + "','" + Trim(Label4.Text) + "','" + Trim(Label5.Text) + "','" + Trim(Label4.Text) + "0101','" + Trim(Label4.Text) + "1231','" + NsaHilo.DataGridView1.Rows(k).Cells(2).Value + "','" + NsaHilo.DataGridView1.Rows(k).Cells(2).Value + "',NULL, 1"
                    Dim cmd1021 As New SqlCommand(sql1021, conx)
                    Rsr21 = cmd1021.ExecuteReader()

                    If Rsr21.Read() Then
                        NsaHilo.DataGridView1.Rows(k).Cells(10).Value = Rsr21(10)
                    Else
                        NsaHilo.DataGridView1.Rows(k).Cells(10).Value = 0
                    End If

                    Rsr21.Close()
                Next
            Else
                If Label6.Text = 3 Then
                    For i = 0 To p - 1
                        If DataGridView1.Rows(i).Cells(0).Value = True Then
                            abrir()
                            Dim i51, i52 As Integer
                            i51 = 0
                            i52 = Guia_hilo.DataGridView1.Rows.Count
                            Dim sql102213 As String = "select * from requerimiento_avios_detalle where id_cabecera ='" + Trim(DataGridView1.Rows(i).Cells(1).Value) + "'"
                            Dim cmd102213 As New SqlCommand(sql102213, conx)
                            Rsr3 = cmd102213.ExecuteReader()
                            While Rsr3.Read()
                                Guia_hilo.DataGridView1.Rows.Add()
                                Guia_hilo.DataGridView1.Rows(i51 + i52).Cells(0).Value = i51 + 1 + i52
                                Guia_hilo.DataGridView1.Rows(i51 + i52).Cells(2).Value = Rsr3(3)
                                Guia_hilo.DataGridView1.Rows(i51 + i52).Cells(3).Value = Rsr3(4)
                                Guia_hilo.DataGridView1.Rows(i51 + i52).Cells(4).Value = Rsr3(8)
                                Guia_hilo.DataGridView1.Rows(i51 + i52).Cells(5).Value = Rsr3(6)
                                Guia_hilo.DataGridView1.Rows(i51 + i52).Cells(7).Value = Rsr3(8)
                                Guia_hilo.DataGridView1.Rows(i51 + i52).Cells(8).Value = Rsr3(6)

                                'Guia_hilo.DataGridView1.Rows(i51 + i52).Cells(11).Value = Rsr3(6)
                                'Guia_hilo.DataGridView1.Rows(i51 + i52).Cells(14).Value = Rsr3(0)
                                'Guia_hilo.DataGridView1.Rows(i51 + i52).Cells(15).Value = Rsr3(1)


                                i51 = i51 + 1
                            End While
                            Rsr3.Close()
                        End If





                    Next
                End If
            End If

        End If

        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Agregar_requerimiento_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        'Dim p As Integer
        'p = DataGridView1.Rows.Count
        'For i = 0 To p - 1
        '    Dim cmd157 As New SqlCommand("update requerimiento_avios_detalle set estado =0 where id_requeminieto_detalle in (@id_requeminieto_detalle)", conx)
        '    cmd157.Parameters.AddWithValue("@id_requeminieto_detalle", Trim(DataGridView1.Rows(i).Cells(11).Value))
        '    cmd157.ExecuteNonQuery()
        'Next
    End Sub
End Class