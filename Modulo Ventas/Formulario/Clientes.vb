Imports System.Data.SqlClient
Public Class Clientes
    Dim dt As New DataTable
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Clientes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fg As New fcliente
        Dim fg1 As New vcliente
        TextBox1.Select()
        fg1.gruc = "01"
        dt = fg.buscar_cliente20(fg1)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(2).Width = 370
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Else
            MsgBox("NO EXISTE INFORMCION PARA MOSTRAR")
        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Ruc" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(2).Width = 370
                DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "R_Social" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(2).Width = 370
                DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar2()
    End Sub
    Sub PO()
        If Orden_Produccion.TextBox20.Text = "SERVICIO" Then
            abrir()
            Dim nu As String
            nu = ""
            Dim Rsr19 As SqlDataReader
            Dim sql101 As String = "SELECT top 1 CONVERT (INT,SUBSTRING( program_3,7,10 )) FROM custom_vianny.dbo.qag0300 WHERE SUBSTRING( program_3,1,6 )= 'SERVIA' AND ccia ='01' order by SUBSTRING( program_3,7,10 ) desc"
            Dim cmd101 As New SqlCommand(sql101, conx)
            Rsr19 = cmd101.ExecuteReader()
            If Rsr19.Read() = False Then
                nu = 1
            Else
                nu = Rsr19(0) + 1
            End If
            Select Case nu.Length
                Case "1" : nu = "000" & "" & nu
                Case "2" : nu = "00" & "" & nu
                Case "3" : nu = "0" & "" & nu
                Case "4" : nu = nu
            End Select
            Orden_Produccion.TextBox3.Text = "SERVIA" + nu
            Rsr19.Close()
        Else
            If Orden_Produccion.TextBox20.Text = "TELA" Then
                abrir()
                Dim nu As String
                nu = ""
                Dim Rsr20 As SqlDataReader
                Dim sql1011 As String = "SELECT top 1 CONVERT (INT,SUBSTRING( program_3,7,10 )) FROM custom_vianny.dbo.qag0300 WHERE SUBSTRING( program_3,1,6 )= 'PROVIA' AND ccia ='01' order by SUBSTRING( program_3,7,10 ) desc"
                Dim cmd1011 As New SqlCommand(sql1011, conx)
                Rsr20 = cmd1011.ExecuteReader()
                If Rsr20.Read() = False Then
                    nu = 1
                Else
                    nu = Rsr20(0) + 1
                End If
                Select Case nu.Length
                    Case "1" : nu = "000" & "" & nu
                    Case "2" : nu = "00" & "" & nu
                    Case "3" : nu = "0" & "" & nu
                    Case "4" : nu = nu
                End Select
                Orden_Produccion.TextBox3.Text = "PROVIA" + nu
                Rsr20.Close()
            Else
                abrir()
                Dim nu As String
                nu = ""
                Dim Rsr21 As SqlDataReader
                Dim sql10111 As String = "SELECT top 1 CONVERT (INT,SUBSTRING( program_3,7,10 )) FROM custom_vianny.dbo.qag0300 WHERE SUBSTRING( program_3,1,6 )= 'MUEVIA' AND ccia ='01' order by SUBSTRING( program_3,7,10 ) desc"
                Dim cmd10111 As New SqlCommand(sql10111, conx)
                Rsr21 = cmd10111.ExecuteReader()
                If Rsr21.Read() = False Then
                    nu = 1
                Else
                    nu = Rsr21(0) + 1
                End If
                Select Case nu.Length
                    Case "1" : nu = "000" & "" & nu
                    Case "2" : nu = "00" & "" & nu
                    Case "3" : nu = "0" & "" & nu
                    Case "4" : nu = nu
                End Select
                Orden_Produccion.TextBox3.Text = "MUEVIA" + nu
                Rsr21.Close()
            End If
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a, conta As Integer
        i = DataGridView1.Rows.Count
        a = 0
        conta = 0

        If TextBox3.Text = 7 Then
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

                        Estado_Cliente.TextBox4.Text = DataGridView1.Rows(a).Cells(1).Value
                        Estado_Cliente.TextBox1.Text = DataGridView1.Rows(a).Cells(2).Value

                    End If
                Next

                Close()
            End If
        Else
            If TextBox3.Text = 18 Then
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


                            HojaCotizacion.TextBox2.Text = DataGridView1.Rows(a).Cells(2).Value

                        End If
                    Next

                    Close()
                End If
            Else
                If TextBox3.Text = 19 Then
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

                                Nota_Pedido.TextBox2.Text = DataGridView1.Rows(a).Cells(1).Value
                                Nota_Pedido.TextBox3.Text = DataGridView1.Rows(a).Cells(2).Value
                                Nota_Pedido.TextBox4.Text = DataGridView1.Rows(a).Cells(3).Value
                            End If
                        Next

                        Close()
                    End If
                Else
                    If TextBox3.Text = 20 Then
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

                                    Guia_remision.TextBox8.Text = DataGridView1.Rows(a).Cells(1).Value
                                    Guia_remision.TextBox9.Text = DataGridView1.Rows(a).Cells(2).Value
                                    Guia_remision.TextBox10.Text = DataGridView1.Rows(a).Cells(3).Value
                                End If
                            Next

                            Close()
                        End If
                    Else
                        If TextBox3.Text = 21 Then
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

                                        Guia_remision.TextBox11.Text = DataGridView1.Rows(a).Cells(1).Value
                                        Guia_remision.TextBox12.Text = DataGridView1.Rows(a).Cells(2).Value
                                        Guia_remision.TextBox13.Text = DataGridView1.Rows(a).Cells(3).Value
                                    End If
                                Next

                                Close()
                            End If
                        Else
                            If TextBox3.Text = 30 Then
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

                                            Nsalida.TextBox8.Text = Trim(DataGridView1.Rows(a).Cells(1).Value).ToString
                                            Nsalida.TextBox10.Text = Trim(DataGridView1.Rows(a).Cells(2).Value).ToString

                                        End If
                                    Next

                                    Close()
                                End If
                            Else
                                If TextBox3.Text = 35 Then
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

                                                NotaIngreso.TextBox8.Text = Trim(DataGridView1.Rows(a).Cells(1).Value).ToString
                                                NotaIngreso.TextBox10.Text = Trim(DataGridView1.Rows(a).Cells(2).Value).ToString

                                            End If
                                        Next

                                        Close()
                                    End If
                                Else
                                    If TextBox3.Text = 88 Then
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

                                                    Reporte_ventasn.TextBox2.Text = Trim(DataGridView1.Rows(a).Cells(1).Value).ToString
                                                    Reporte_ventasn.TextBox1.Text = Trim(DataGridView1.Rows(a).Cells(2).Value).ToString

                                                End If
                                            Next

                                            Close()
                                        End If
                                    Else
                                        If TextBox3.Text = 300 Then
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

                                                        Guia_hilo.TextBox8.Text = DataGridView1.Rows(a).Cells(1).Value
                                                        Guia_hilo.TextBox9.Text = DataGridView1.Rows(a).Cells(2).Value
                                                        Guia_hilo.TextBox10.Text = DataGridView1.Rows(a).Cells(3).Value

                                                    End If
                                                Next

                                                Close()


                                            End If
                                        Else
                                            If TextBox3.Text = 400 Then
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

                                                            NsaHilo.TextBox8.Text = DataGridView1.Rows(a).Cells(1).Value
                                                            NsaHilo.TextBox10.Text = DataGridView1.Rows(a).Cells(2).Value
                                                            'NsaHilo.TextBox9.Text = DataGridView1.Rows(a).Cells(3).Value

                                                        End If
                                                    Next

                                                    Close()
                                                End If
                                            Else
                                                If TextBox3.Text = 350 Then
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

                                                                NiaHilo.TextBox8.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                NiaHilo.TextBox10.Text = DataGridView1.Rows(a).Cells(2).Value
                                                                NiaHilo.TextBox9.Text = DataGridView1.Rows(a).Cells(3).Value

                                                            End If
                                                        Next

                                                        Close()
                                                    End If
                                                Else
                                                    If TextBox3.Text = 333 Then
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

                                                                    Ventas_Cliente.TextBox1.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                    Ventas_Cliente.TextBox2.Text = DataGridView1.Rows(a).Cells(2).Value


                                                                End If
                                                            Next

                                                            Close()
                                                        End If
                                                    Else
                                                        If TextBox3.Text = 39 Then
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

                                                                        Otejeduria.TextBox3.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                        Otejeduria.TextBox4.Text = DataGridView1.Rows(a).Cells(2).Value


                                                                    End If
                                                                Next

                                                                Close()
                                                            End If
                                                        Else
                                                            If TextBox3.Text = 40 Then
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

                                                                            Otejeduria.TextBox8.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                            Otejeduria.TextBox9.Text = DataGridView1.Rows(a).Cells(2).Value


                                                                        End If
                                                                    Next

                                                                    Close()
                                                                End If
                                                            Else
                                                                If TextBox3.Text = 335 Then
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

                                                                                Orden_Produccion.TextBox9.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                                Orden_Produccion.TextBox2.Text = DataGridView1.Rows(a).Cells(2).Value

                                                                            End If
                                                                        Next

                                                                        Close()
                                                                    End If
                                                                Else
                                                                    If TextBox3.Text = 339 Then
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
                                                                                    Dim o As Integer
                                                                                    o = Label3.Text
                                                                                    Otejeduria.DataGridView2.Rows(o).Cells(3).Value = Trim(DataGridView1.Rows(a).Cells(1).Value)
                                                                                    Otejeduria.DataGridView2.CurrentCell = Otejeduria.DataGridView2.Item(2, o)


                                                                                End If
                                                                            Next

                                                                            Close()
                                                                        End If
                                                                    Else
                                                                        If TextBox3.Text = 3363 Then
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

                                                                                        Requerimiento__comercial.TextBox1.Text = Trim(DataGridView1.Rows(a).Cells(2).Value)


                                                                                    End If
                                                                                Next

                                                                                Close()
                                                                            End If
                                                                        Else
                                                                            If TextBox3.Text = 111 Then
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
                                                                                            FICHA_TECNICA.TextBox30.Text = Trim(DataGridView1.Rows(a).Cells(1).Value)
                                                                                            FICHA_TECNICA.TextBox3.Text = Trim(DataGridView1.Rows(a).Cells(2).Value)


                                                                                        End If
                                                                                    Next

                                                                                    Close()
                                                                                End If
                                                                            Else
                                                                                If TextBox3.Text = 3555 Then
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

                                                                                                Ocs.TextBox3.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                                                Ocs.TextBox4.Text = DataGridView1.Rows(a).Cells(2).Value

                                                                                            End If
                                                                                        Next

                                                                                        Close()
                                                                                    End If
                                                                                Else
                                                                                    If TextBox3.Text = 35556 Then
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

                                                                                                    os.TextBox3.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                                                    os.TextBox4.Text = DataGridView1.Rows(a).Cells(2).Value

                                                                                                End If
                                                                                            Next

                                                                                            Close()
                                                                                        End If
                                                                                    Else
                                                                                        If TextBox3.Text = "800" Then
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

                                                                                                        Hoja_Reclamos.TextBox2.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                                                        Hoja_Reclamos.TextBox3.Text = DataGridView1.Rows(a).Cells(2).Value

                                                                                                    End If
                                                                                                Next

                                                                                                Close()
                                                                                            End If
                                                                                        Else
                                                                                            If TextBox3.Text = "35080" Then
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

                                                                                                            Nsa_Pt.TextBox8.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                                                            Nsa_Pt.TextBox10.Text = DataGridView1.Rows(a).Cells(2).Value

                                                                                                        End If
                                                                                                    Next

                                                                                                    Close()
                                                                                                End If
                                                                                            Else
                                                                                                If TextBox3.Text = "36080" Then
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

                                                                                                                Nia_Pt.TextBox8.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                                                                Nia_Pt.TextBox10.Text = DataGridView1.Rows(a).Cells(2).Value

                                                                                                            End If
                                                                                                        Next

                                                                                                        Close()
                                                                                                    End If
                                                                                                Else
                                                                                                    If TextBox3.Text = "39080" Then
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

                                                                                                                    Guia_remision_prenda.TextBox8.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                                                                    Guia_remision_prenda.TextBox9.Text = DataGridView1.Rows(a).Cells(2).Value
                                                                                                                    Guia_remision_prenda.TextBox10.Text = DataGridView1.Rows(a).Cells(3).Value
                                                                                                                End If
                                                                                                            Next

                                                                                                            Close()
                                                                                                        End If
                                                                                                    Else
                                                                                                        If TextBox3.Text = "22" Then
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

                                                                                                                        RegisroInconvenientes.TextBox1.Text = DataGridView1.Rows(a).Cells(1).Value
                                                                                                                        RegisroInconvenientes.TextBox2.Text = DataGridView1.Rows(a).Cells(2).Value

                                                                                                                    End If
                                                                                                                Next

                                                                                                                Close()
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                                End If
                                                                                            End If
                                                                                    End If
                                                                                    End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If

                                                End If
                                            End If
                                        End If


                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub
End Class