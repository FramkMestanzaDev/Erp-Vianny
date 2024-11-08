Imports System.Data.SqlClient
Public Class Productos_Vianny
    Public conx As SqlConnection
    Public comando As SqlCommand
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Tipo_Venta.Show()
    End Sub
    Dim dt As New DataTable
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(gk.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PRODUCTO" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                Dim K As Integer
                Dim jk As Double
                K = DataGridView1.Rows.Count
                If k > 0 Then
                    For i = 0 To k - 1
                        jk = jk + DataGridView1.Rows(i).Cells(6).Value
                        If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                            DataGridView1.Rows(i).Cells(11).Value = "SI"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                            DataGridView1.Rows(i).Cells(11).Value = "NO"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                            DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(14).Value = 1 Then

                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                        End If
                    Next
                    TextBox4.Text = jk
                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar_PARTIDA()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(gk.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PARTIDA" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                Dim K As Integer
                Dim jk As Double
                K = DataGridView1.Rows.Count
                If k > 0 Then
                    For i = 0 To k - 1
                        jk = jk + DataGridView1.Rows(i).Cells(6).Value
                        If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                            DataGridView1.Rows(i).Cells(11).Value = "SI"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                            DataGridView1.Rows(i).Cells(11).Value = "NO"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                            DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(14).Value = 1 Then

                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                        End If
                    Next
                    TextBox4.Text = jk
                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(gk1.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PRODUCTO" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                Dim K As Integer
                Dim jk As Double
                K = DataGridView1.Rows.Count
                If k > 0 Then
                    For i = 0 To k - 1
                        jk = jk + DataGridView1.Rows(i).Cells(6).Value
                        If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                            DataGridView1.Rows(i).Cells(11).Value = "SI"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                            DataGridView1.Rows(i).Cells(11).Value = "NO"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                            DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(14).Value = 1 Then

                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                        End If
                    Next
                    TextBox4.Text = jk
                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar_tercera()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(gj3.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PRODUCTO" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                Dim K As Integer
                Dim jk As Double
                K = DataGridView1.Rows.Count
                If K > 0 Then
                    For i = 0 To K - 1
                        jk = jk + DataGridView1.Rows(i).Cells(6).Value
                        If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                            DataGridView1.Rows(i).Cells(11).Value = "SI"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                            DataGridView1.Rows(i).Cells(11).Value = "NO"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                            DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(14).Value = 1 Then

                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                        End If
                    Next
                    TextBox4.Text = jk
                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar_partida2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(gk1.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PARTIDA" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                Dim K As Integer
                Dim jk As Double
                K = DataGridView1.Rows.Count
                If k > 0 Then
                    For i = 0 To k - 1
                        jk = jk + DataGridView1.Rows(i).Cells(6).Value
                        If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                            DataGridView1.Rows(i).Cells(11).Value = "SI"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                            DataGridView1.Rows(i).Cells(11).Value = "NO"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                            DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(14).Value = 1 Then

                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                        End If
                    Next
                    TextBox4.Text = jk
                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar3()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(gk2.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PRODUCTO" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                Dim K As Integer
                Dim jk As Double
                K = DataGridView1.Rows.Count
                If k > 0 Then
                    For i = 0 To k - 1
                        jk = jk + DataGridView1.Rows(i).Cells(6).Value
                        If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                            DataGridView1.Rows(i).Cells(11).Value = "SI"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                            DataGridView1.Rows(i).Cells(11).Value = "NO"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                            DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(14).Value = 1 Then

                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                        End If
                    Next
                    TextBox4.Text = jk
                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub buscar_PARTIDA3()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(gk2.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PARTIDA" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                Dim K As Integer
                Dim jk As Double
                K = DataGridView1.Rows.Count
                If k > 0 Then
                    For i = 0 To k - 1
                        jk = jk + DataGridView1.Rows(i).Cells(6).Value
                        If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                            DataGridView1.Rows(i).Cells(11).Value = "SI"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                            DataGridView1.Rows(i).Cells(11).Value = "NO"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                            DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(14).Value = 1 Then

                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                        End If
                    Next
                    TextBox4.Text = jk
                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar_PARTIDA_tercera()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(gj3.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PARTIDA" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
                Dim K As Integer
                Dim jk As Double
                K = DataGridView1.Rows.Count
                If K > 0 Then
                    For i = 0 To K - 1
                        jk = jk + DataGridView1.Rows(i).Cells(6).Value
                        If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                            DataGridView1.Rows(i).Cells(11).Value = "SI"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                            DataGridView1.Rows(i).Cells(11).Value = "NO"
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                            DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                            DataGridView1.Columns(11).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(14).Value = 1 Then

                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                        End If
                    Next
                    TextBox4.Text = jk
                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Dim Rs4 As SqlDataReader
    Private Sub Productos_Vianny_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TG As New FSTOCK_PAR
        Dim HJ As New VSTOCK_PAR
        Dim k As Integer
        Dim jk As Double
        RadioButton7.Checked = True
        HJ.gano = Nota_Pedido.Label25.Text
        HJ.gCCIA = Nota_Pedido.Label26.Text
        TextBox3.Text = ""
        If Trim(Nota_Pedido.Label26.Text) = "01" Or Trim(Nota_Pedido.Label26.Text) = "03" Then

            gk = TG.CaeSoft_ReporteStockFisico_COMERCIAL(HJ)

        Else
            gk = TG.CaeSoft_ReporteStockFisico_COMERCIAL_GRAUS(HJ)
            RadioButton1.Visible = False
            RadioButton8.Visible = False
            RadioButton7.Text = "HILO - TELA GRAUS"

        End If

        DataGridView1.DataSource = gk
        k = DataGridView1.Rows.Count
        If k > 0 Then
            For i = 0 To k - 1
                jk = jk + DataGridView1.Rows(i).Cells(6).Value
                If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                    DataGridView1.Rows(i).Cells(11).Value = "SI"
                    DataGridView1.Columns(11).ReadOnly = True
                End If
                If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                    DataGridView1.Rows(i).Cells(11).Value = "NO"
                    DataGridView1.Columns(11).ReadOnly = True
                End If
                If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                    DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                    DataGridView1.Columns(11).ReadOnly = True
                End If
                If DataGridView1.Rows(i).Cells(14).Value = 1 Then

                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                End If
            Next
            TextBox4.Text = jk
            DataGridView1.Columns(1).Width = 145
            DataGridView1.Columns(2).Width = 146
            DataGridView1.Columns(3).Width = 340
            DataGridView1.Columns(5).Width = 75
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Width = 45
            DataGridView1.Columns(8).Width = 80

            DataGridView1.Columns(9).Width = 80
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(14).Visible = False
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim i, a As Integer
        i = DataGridView1.Rows.Count

        Dim u As Integer
        For a = 0 To i - 1

            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                Nota_Pedido.DataGridView1.Rows.Add(1)
                u = Nota_Pedido.DataGridView1.Rows.Count
                If u = 1 Then



                    Nota_Pedido.DataGridView1.Rows(0).Cells(0).Value = u
                    If RadioButton7.Checked = True Or RadioButton7.Checked = True Then
                        Nota_Pedido.DataGridView1.Rows(u - 1).Cells(1).Value = "PRIMERA"
                    Else
                        If RadioButton8.Checked = True Then
                            Nota_Pedido.DataGridView1.Rows(u - 1).Cells(1).Value = "SEGUNDA"
                        Else
                            Nota_Pedido.DataGridView1.Rows(u - 1).Cells(1).Value = "TERCERA"
                        End If

                    End If
                    Nota_Pedido.DataGridView1.Rows(0).Cells(14).Value = Mid(DataGridView1.Rows(a).Cells(1).Value, 2, 2)
                    Nota_Pedido.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(a).Cells(2).Value
                        Nota_Pedido.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(a).Cells(3).Value
                        Nota_Pedido.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(a).Cells(4).Value
                    'Nota_Pedido.DataGridView1.Rows(0).Cells(5).Value = DataGridView1.Rows(a).Cells(9).Value
                    'Nota_Pedido.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(a).Cells(8).Value

                    Nota_Pedido.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(a).Cells(7).Value
                        Nota_Pedido.DataGridView1.Rows(0).Cells(8).Value = DataGridView1.Rows(a).Cells(6).Value - DataGridView1.Rows(a).Cells(15).Value
                        Nota_Pedido.DataGridView1.Rows(0).Cells(13).Value = DataGridView1.Rows(a).Cells(6).Value - DataGridView1.Rows(a).Cells(15).Value


                    Else
                        If u > 1 Then


                        'If Trim(Nota_Pedido.DataGridView1.Rows(0).Cells(14).Value) = Mid(DataGridView1.Rows(a).Cells(1).Value, 2, 2) Then
                        Nota_Pedido.DataGridView1.Rows(u - 1).Cells(0).Value = u

                        If RadioButton7.Checked = True Or RadioButton7.Checked = True Then
                            Nota_Pedido.DataGridView1.Rows(u - 1).Cells(1).Value = "PRIMERA"
                        Else
                            If RadioButton8.Checked = True Then
                                Nota_Pedido.DataGridView1.Rows(u - 1).Cells(1).Value = "SEGUNDA"
                            Else
                                Nota_Pedido.DataGridView1.Rows(u - 1).Cells(1).Value = "TERCERA"
                                End If

                            End If
                            Nota_Pedido.DataGridView1.Rows(u - 1).Cells(14).Value = Mid(DataGridView1.Rows(a).Cells(1).Value, 2, 2)
                            Nota_Pedido.DataGridView1.Rows(u - 1).Cells(2).Value = DataGridView1.Rows(a).Cells(2).Value
                            Nota_Pedido.DataGridView1.Rows(u - 1).Cells(3).Value = DataGridView1.Rows(a).Cells(3).Value
                            Nota_Pedido.DataGridView1.Rows(u - 1).Cells(4).Value = DataGridView1.Rows(a).Cells(4).Value
                        'Nota_Pedido.DataGridView1.Rows(u - 1).Cells(5).Value = DataGridView1.Rows(a).Cells(9).Value
                        'Nota_Pedido.DataGridView1.Rows(u - 1).Cells(6).Value = DataGridView1.Rows(a).Cells(8).Value

                        Nota_Pedido.DataGridView1.Rows(u - 1).Cells(7).Value = DataGridView1.Rows(a).Cells(7).Value
                            Nota_Pedido.DataGridView1.Rows(u - 1).Cells(8).Value = DataGridView1.Rows(a).Cells(6).Value - DataGridView1.Rows(a).Cells(15).Value
                            Nota_Pedido.DataGridView1.Rows(u - 1).Cells(13).Value = DataGridView1.Rows(a).Cells(6).Value - DataGridView1.Rows(a).Cells(15).Value

                        'Else
                        '    MsgBox("NO SE PUEDE SELECCIONAR PRODUCTOS DE DIFERENTES ALMACENES")
                        '    Nota_Pedido.DataGridView1.Rows.Remove(Nota_Pedido.DataGridView1.Rows(u - 1))
                        'End If
                    End If



                End If
            End If
        Next


        Close()

    End Sub
    Dim gk, gk1, gk2, gj3, gk50 As New DataTable
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim TG As New FSTOCK_PAR
        Dim HJ As New VSTOCK_PAR
        Dim k As Integer
        Dim jk As Double
        TextBox1.Text = ""
        TextBox3.Text = ""
        DataGridView1.DataSource = ""
        Try
            If RadioButton7.Checked = True Then

                HJ.gano = Label4.Text
                HJ.gCCIA = Label5.Text

                If Trim(Label5.Text) = "01" Or Trim(Label5.Text) = "03" Then

                    gk50 = TG.CaeSoft_ReporteStockFisico_COMERCIAL(HJ)
                Else

                    gk50 = TG.CaeSoft_ReporteStockFisico_COMERCIAL_GRAUS(HJ)
                End If
                DataGridView1.DataSource = gk50

            Else
                If RadioButton8.Checked = True Then
                    HJ.gano = Label4.Text
                    HJ.gCCIA = Label5.Text
                    gk1 = TG.CaeSoft_ReporteStockFisico_COMERCIA2L(HJ)
                    DataGridView1.DataSource = gk1
                Else
                    If RadioButton1.Checked = True Then
                        HJ.gALMACEN = "03"
                        HJ.gCCIA = Label5.Text
                        HJ.gMODAL = "02"
                        HJ.gano = Label4.Text
                        gk2 = TG.CaeSoft_ReporteStockFisico_3737(HJ)
                        DataGridView1.DataSource = gk2
                    Else
                        If RadioButton2.Checked = True Then
                            HJ.gano = Label4.Text
                            HJ.gCCIA = Label5.Text
                            If Label5.Text = "01" And Label5.Text = "03" Then
                                gj3 = TG.CaeSoft_ReporteStockFisico_tercera(HJ)
                                DataGridView1.DataSource = gj3
                            End If
                        End If

                    End If
                End If
            End If
            k = DataGridView1.Rows.Count
            If k > 0 Then
                For i = 0 To k - 1
                    jk = jk + Convert.ToDouble(DataGridView1.Rows(i).Cells(6).Value)
                    If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                        DataGridView1.Rows(i).Cells(11).Value = "SI"
                        DataGridView1.Columns(11).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                        DataGridView1.Rows(i).Cells(11).Value = "NO"
                        DataGridView1.Columns(11).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                        DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                        DataGridView1.Columns(11).ReadOnly = True
                    End If
                    If Trim(DataGridView1.Rows(i).Cells(14).Value) = "1" Then

                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                    End If
                Next
                TextBox4.Text = jk
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
            End If
        Catch ex As Exception
            MsgBox("NO HAY STOCK INICIAL DEL AÑO EN CURSO")
        End Try




    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString

            TextBox4.Text = DataGridView1.Item(6, num1).Value()

        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If RadioButton7.Checked = True Then
            buscar_PARTIDA()

        Else
            If RadioButton8.Checked = True Then
                buscar_partida2()
            Else
                If RadioButton1.Checked = True Then
                    buscar_PARTIDA3()
                Else
                    If RadioButton2.Checked = True Then
                        buscar_PARTIDA_tercera()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If RadioButton7.Checked = True Then
            buscar()

        Else
            If RadioButton8.Checked = True Then
                buscar2()
            Else
                If RadioButton1.Checked = True Then
                    buscar3()
                Else
                    If RadioButton2.Checked = True Then
                        buscar_tercera()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs)

    End Sub
End Class