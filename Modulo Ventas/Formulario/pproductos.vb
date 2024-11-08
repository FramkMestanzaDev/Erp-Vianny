Imports System.Data.SqlClient
Public Class pproductos
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand
    Public Property Padre As Form_Cotizacion
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim fecha1 As Date
    Dim bs As New BindingSource()

    Dim da As New DataTable
    Dim filtroBase As String = ""
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim textoBusqueda As String = TextBox2.Text.Trim()

        ' Si el texto contiene el símbolo %, dividimos en partes
        If textoBusqueda.Contains("%") Then
            Dim partes As String() = textoBusqueda.Split(New Char() {"%"c}, StringSplitOptions.None)

            ' Llamamos a la función AplicarFiltro con las primeras 4 partes (si existen)
            AplicarFiltro(partes)
        Else
            ' Si no se encuentra el %, simplemente aplicar el filtro en tiempo real
            AplicarFiltro(New String() {textoBusqueda})
        End If
    End Sub
    Private Sub AplicarFiltro(partes As String())
        ' Verifica si el DataSource es un BindingSource
        If TypeOf DataGridView1.DataSource Is BindingSource Then
            Dim bs As BindingSource = CType(DataGridView1.DataSource, BindingSource)
            If bs IsNot Nothing AndAlso bs.DataSource IsNot Nothing Then
                Dim dt As DataTable = CType(bs.DataSource, DataTable)

                If dt IsNot Nothing Then
                    ' Escapamos caracteres especiales en cada parte
                    For i As Integer = 0 To partes.Length - 1
                        partes(i) = partes(i).Replace("'", "''").Trim()
                    Next

                    ' Construimos el filtro para aplicar hasta 4 partes
                    Dim filtro As String = ""
                    If partes.Length > 0 Then filtro &= String.Format("Nomb_17 LIKE '%{0}%'", partes(0))
                    If partes.Length > 1 Then filtro &= String.Format(" AND Nomb_17 LIKE '%{0}%'", partes(1))
                    If partes.Length > 2 Then filtro &= String.Format(" AND Nomb_17 LIKE '%{0}%'", partes(2))
                    If partes.Length > 3 Then filtro &= String.Format(" AND Nomb_17 LIKE '%{0}%'", partes(3))

                    ' Aplicamos el filtro
                    dt.DefaultView.RowFilter = filtro
                End If
            End If

        End If
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim i, I2, SUMA As Integer
        If Label3.Text = 1 Then
            i = DataGridView1.Rows.Count
            For O = 0 To i - 1
                If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                    NotaIngreso.DataGridView1.Rows.Add(1)
                    I2 = NotaIngreso.DataGridView1.Rows.Count

                    If I2 = 1 Then


                        NotaIngreso.DataGridView1.Rows(0).Cells(0).Value = 1

                        NotaIngreso.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(O).Cells(2).Value
                        NotaIngreso.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(O).Cells(4).Value
                        NotaIngreso.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(O).Cells(5).Value

                    Else
                        If I2 > 1 Then

                            SUMA = NotaIngreso.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(0).Value = SUMA

                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(O).Cells(2).Value
                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(O).Cells(4).Value
                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(O).Cells(5).Value



                        End If
                    End If
                End If
            Next
        Else
            If Label3.Text = 2 Then
                i = DataGridView1.Rows.Count
                For O = 0 To i - 1
                    If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                        NiaHilo.DataGridView1.Rows.Add(1)
                        I2 = NiaHilo.DataGridView1.Rows.Count

                        If I2 = 1 Then


                            NiaHilo.DataGridView1.Rows(0).Cells(0).Value = 1

                            NiaHilo.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(O).Cells(2).Value
                            NiaHilo.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(O).Cells(4).Value
                            NiaHilo.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(O).Cells(5).Value

                        Else
                            If I2 > 1 Then

                                SUMA = NiaHilo.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                                NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(0).Value = SUMA

                                NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(O).Cells(2).Value
                                NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(O).Cells(4).Value
                                NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(7).Value = DataGridView1.Rows(O).Cells(5).Value



                            End If
                        End If
                    End If
                Next
            Else
                If Label3.Text = 3 Then
                    i = DataGridView1.Rows.Count


                    Dim a, conta As Integer

                    I2 = os.DataGridView1.Rows.Count


                    conta = 0



                    For A = 0 To i - 1
                        If Me.DataGridView1.Rows(A).Cells(0).Value = True Then
                            conta = conta + 1
                        End If
                    Next
                    If conta > 1 Then
                        MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                    Else
                        For O = 0 To i - 1
                            If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                                os.DataGridView1.Rows(Label5.Text).Cells(0).Value = Convert.ToInt32(Label5.Text) + 1

                                os.DataGridView1.Rows(Label5.Text).Cells(4).Value = DataGridView1.Rows(O).Cells(2).Value
                                os.DataGridView1.Rows(Label5.Text).Cells(5).Value = DataGridView1.Rows(O).Cells(4).Value
                                os.DataGridView1.Rows(Label5.Text).Cells(6).Value = DataGridView1.Rows(O).Cells(5).Value

                            End If
                        Next
                        'If I2 = 1 Then


                        '    os.DataGridView1.Rows(0).Cells(0).Value = 1

                        '    os.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(O).Cells(2).Value
                        '    os.DataGridView1.Rows(0).Cells(5).Value = DataGridView1.Rows(O).Cells(4).Value
                        '    os.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(O).Cells(5).Value

                        'Else
                        '    If I2 > 1 Then

                        '        SUMA = os.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                        '        os.DataGridView1.Rows(SUMA - 1).Cells(0).Value = SUMA

                        '        os.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(O).Cells(2).Value
                        '        os.DataGridView1.Rows(SUMA - 1).Cells(5).Value = DataGridView1.Rows(O).Cells(4).Value
                        '        os.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(O).Cells(5).Value



                        '    End If
                        'End If
                        Me.Close()
                    End If
                Else
                    If Label3.Text = 4 Then
                        i = DataGridView1.Rows.Count
                        Dim a, conta As Integer
                        I2 = Ocs.DataGridView1.Rows.Count
                        conta = 0
                        For a = 0 To i - 1
                            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                conta = conta + 1
                            End If
                        Next
                        If conta > 1 Then
                            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                        Else
                            For O = 0 To i - 1
                                If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                                    Ocs.DataGridView1.Rows(Label5.Text).Cells(0).Value = Convert.ToInt32(Label5.Text) + 1

                                    Ocs.DataGridView1.Rows(Label5.Text).Cells(5).Value = DataGridView1.Rows(O).Cells(2).Value
                                    Ocs.DataGridView1.Rows(Label5.Text).Cells(6).Value = DataGridView1.Rows(O).Cells(4).Value
                                    Ocs.DataGridView1.Rows(Label5.Text).Cells(10).Value = DataGridView1.Rows(O).Cells(5).Value

                                End If
                            Next
                        End If
                    Else
                        If Label3.Text = 5 Then
                            i = DataGridView1.Rows.Count


                            Dim a, conta As Integer

                            I2 = Ocs.DataGridView1.Rows.Count


                            conta = 0



                            For a = 0 To i - 1
                                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                    conta = conta + 1
                                End If
                            Next
                            If conta > 1 Then
                                MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                            Else
                                For O = 0 To i - 1
                                    If Me.DataGridView1.Rows(O).Cells(0).Value = True Then


                                        Formato_solic_avios.TextBox3.Text = DataGridView1.Rows(O).Cells(2).Value
                                        Formato_solic_avios.TextBox3.Focus()
                                        SendKeys.Send("{ENTER}")
                                        Me.Close()
                                    End If
                                Next
                            End If
                        Else
                            If Label3.Text = 6 Then
                                i = DataGridView1.Rows.Count


                                Dim I3 As Integer

                                For O = 0 To i - 1
                                    If Me.DataGridView1.Rows(O).Cells(0).Value = True Then

                                        ProductosLinea.DataGridView1.Rows.Add()
                                        I3 = ProductosLinea.DataGridView1.Rows.Count

                                        ProductosLinea.DataGridView1.Rows(I3 - 1).Cells(0).Value = TextBox1.Text
                                        ProductosLinea.DataGridView1.Rows(I3 - 1).Cells(1).Value = DataGridView1.Rows(O).Cells(2).Value
                                        ProductosLinea.DataGridView1.Rows(I3 - 1).Cells(2).Value = DataGridView1.Rows(O).Cells(4).Value
                                    End If
                                Next
                                Me.Close()
                            Else
                                If Label3.Text = 7 Then
                                    i = DataGridView1.Rows.Count
                                    Dim a, conta As Integer
                                    I2 = Ocs.DataGridView1.Rows.Count
                                    conta = 0
                                    For a = 0 To i - 1
                                        If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                            conta = conta + 1
                                        End If
                                    Next
                                    If conta > 1 Then
                                        MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                    Else
                                        For O = 0 To i - 1
                                            If Me.DataGridView1.Rows(O).Cells(0).Value = True Then

                                                Od_Udp.TextBox24.Text = DataGridView1.Rows(O).Cells(2).Value
                                                Od_Udp.TextBox25.Text = DataGridView1.Rows(O).Cells(4).Value
                                                Od_Udp.TextBox6.Text = DataGridView1.Rows(O).Cells(4).Value
                                            End If
                                        Next
                                        Me.Close()
                                    End If
                                Else
                                    If Label3.Text = 8 Then
                                        i = DataGridView1.Rows.Count
                                        Dim a, conta As Integer
                                        I2 = Ocs.DataGridView1.Rows.Count
                                        conta = 0
                                        For a = 0 To i - 1
                                            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                                conta = conta + 1
                                            End If
                                        Next
                                        If conta > 1 Then
                                            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                        Else
                                            For O = 0 To i - 1
                                                If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                                                    Od_Udp.TextBox26.Text = DataGridView1.Rows(O).Cells(2).Value
                                                    Od_Udp.TextBox27.Text = DataGridView1.Rows(O).Cells(4).Value
                                                End If
                                            Next
                                            Me.Close()
                                        End If
                                    Else
                                        If Label3.Text = 9 Then
                                            i = DataGridView1.Rows.Count
                                            Dim a, conta As Integer
                                            I2 = Ocs.DataGridView1.Rows.Count
                                            conta = 0
                                            For a = 0 To i - 1
                                                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                                    conta = conta + 1
                                                End If
                                            Next
                                            If conta > 1 Then
                                                MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                            Else
                                                For O = 0 To i - 1
                                                    If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                                                        Od_Udp.TextBox28.Text = DataGridView1.Rows(O).Cells(2).Value
                                                        Od_Udp.TextBox29.Text = DataGridView1.Rows(O).Cells(4).Value
                                                    End If
                                                Next
                                                Me.Close()
                                            End If
                                        Else
                                            If Label3.Text = 10 Then
                                                i = DataGridView1.Rows.Count
                                                Dim a, conta As Integer
                                                I2 = Ocs.DataGridView1.Rows.Count
                                                conta = 0
                                                For a = 0 To i - 1
                                                    If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                                        conta = conta + 1
                                                    End If
                                                Next
                                                If conta > 1 Then
                                                    MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                                Else
                                                    For O = 0 To i - 1
                                                        If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                                                            Od_Udp.TextBox30.Text = DataGridView1.Rows(O).Cells(2).Value
                                                            Od_Udp.TextBox31.Text = DataGridView1.Rows(O).Cells(4).Value
                                                        End If
                                                    Next
                                                    Me.Close()
                                                End If
                                            Else
                                                If Label3.Text = 11 Then
                                                    i = DataGridView1.Rows.Count
                                                    Dim a, conta As Integer
                                                    I2 = Ocs.DataGridView1.Rows.Count
                                                    conta = 0
                                                    For a = 0 To i - 1
                                                        If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                                            conta = conta + 1
                                                        End If
                                                    Next
                                                    If conta > 1 Then
                                                        MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                                    Else
                                                        For O = 0 To i - 1
                                                            If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                                                                OD.TextBox95.Text = Trim(DataGridView1.Rows(O).Cells(2).Value)
                                                                OD.TextBox75.Text = Trim(DataGridView1.Rows(O).Cells(4).Value)

                                                            End If
                                                        Next
                                                        Me.Close()
                                                    End If
                                                Else
                                                    If Label3.Text = 12 Then
                                                        i = DataGridView1.Rows.Count
                                                        Dim a, conta As Integer


                                                        conta = 0
                                                        For a = 0 To i - 1
                                                            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                                                conta = conta + 1
                                                            End If
                                                        Next
                                                        If conta > 1 Then
                                                            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                                        Else
                                                            For O = 0 To i - 1
                                                                If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                                                                    Matriz_Avios.DataGridView1.Rows(Label5.Text).Cells(6).Value = DataGridView1.Rows(O).Cells(2).Value
                                                                    Matriz_Avios.DataGridView1.Rows(Label5.Text).Cells(7).Value = DataGridView1.Rows(O).Cells(4).Value
                                                                End If
                                                            Next

                                                        End If
                                                        Me.Close()
                                                    Else
                                                        If Label3.Text = 13 Then
                                                            i = DataGridView1.Rows.Count
                                                            Dim a, conta As Integer


                                                            conta = 0
                                                            For a = 0 To i - 1
                                                                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                                                    conta = conta + 1
                                                                End If
                                                            Next
                                                            If conta > 1 Then
                                                                MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                                            Else
                                                                For O = 0 To i - 1
                                                                    If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                                                                        DetalleMatrizAvios.DataGridView1.Rows(Label5.Text).Cells(3).Value = DataGridView1.Rows(O).Cells(2).Value
                                                                        DetalleMatrizAvios.DataGridView1.Rows(Label5.Text).Cells(4).Value = DataGridView1.Rows(O).Cells(4).Value
                                                                    End If
                                                                Next

                                                            End If
                                                            Me.Close()
                                                        Else
                                                            If Label3.Text = 14 Then
                                                                i = DataGridView1.Rows.Count
                                                                Dim a, conta As Integer


                                                                conta = 0
                                                                For a = 0 To i - 1
                                                                    If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                                                        conta = conta + 1
                                                                    End If
                                                                Next
                                                                If conta > 1 Then
                                                                    MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                                                Else
                                                                    For O = 0 To i - 1
                                                                        If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                                                                            MatrizAviosOd.DataGridView1.Rows(Label5.Text).Cells(6).Value = DataGridView1.Rows(O).Cells(2).Value.ToString().Trim()
                                                                            MatrizAviosOd.DataGridView1.Rows(Label5.Text).Cells(7).Value = DataGridView1.Rows(O).Cells(4).Value.ToString().Trim()
                                                                        End If
                                                                    Next

                                                                End If
                                                                Me.Close()
                                                            Else
                                                                If Label3.Text = 15 Then
                                                                    i = DataGridView1.Rows.Count
                                                                    Dim a, conta As Integer


                                                                    conta = 0
                                                                    For a = 0 To i - 1
                                                                        If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                                                            conta = conta + 1
                                                                        End If
                                                                    Next
                                                                    If conta > 1 Then
                                                                        MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                                                    Else
                                                                        For O = 0 To i - 1
                                                                            If Me.DataGridView1.Rows(O).Cells(0).Value = True Then
                                                                                Orden_Compra_Ac.TextBox3.Text = DataGridView1.Rows(O).Cells(2).Value.ToString().Trim()
                                                                                'MatrizAviosOd.DataGridView1.Rows(Label5.Text).Cells(7).Value = DataGridView1.Rows(O).Cells(4).Value
                                                                            End If
                                                                        Next

                                                                    End If
                                                                    Me.Close()
                                                                Else
                                                                    If Label3.Text = "16" Then
                                                                        i = DataGridView1.Rows.Count
                                                                        Dim a, conta As Integer
                                                                        conta = 0
                                                                        For a = 0 To i - 1
                                                                            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                                                                conta = conta + 1
                                                                            End If
                                                                        Next
                                                                        If conta > 1 Then
                                                                            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                                                        Else
                                                                            For O = 0 To i - 1
                                                                                If Me.DataGridView1.Rows(O).Cells(0).Value = True Then

                                                                                    Padre.TextBox1.Text = DataGridView1.Rows(O).Cells(2).Value.ToString().Trim()
                                                                                    Padre.TextBox2.Text = DataGridView1.Rows(O).Cells(4).Value.ToString().Trim()
                                                                                    Padre.TextBox4.Text = DataGridView1.Rows(O).Cells(5).Value.ToString().Trim()
                                                                                End If
                                                                            Next
                                                                        End If
                                                                        Me.Close()
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
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Private Sub pproductos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = ""
        da.Clear()
        abrir()
        llenar_combo_box2()
        Dim cmd As New SqlDataAdapter("EXEC Sp_AyudaStockxAlmacen '" + CCIA.Text + "' , '" + ALMACEN.Text + "', NULL, '" + ANO.Text + "', '" + Replace(FECHA.Text, "-", "") + "', NULL", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(2).Width = 180
            DataGridView1.Columns(4).Width = 450
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(11).Visible = True
            DataGridView1.Columns(13).Frozen = True
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            DataGridView1.Columns(4).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            DataGridView1.Columns(6).Frozen = True
            DataGridView1.Columns(7).Frozen = True
            DataGridView1.Columns(8).Frozen = True
            DataGridView1.Columns(9).Frozen = True
            DataGridView1.Columns(10).Frozen = True
            DataGridView1.Columns(11).Frozen = True
            DataGridView1.Columns(12).Frozen = True
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).ReadOnly = True
            DataGridView1.Columns(4).ReadOnly = True
            DataGridView1.Columns(5).ReadOnly = True
        Else
            DataGridView1.DataSource = Nothing
        End If
        abrir()
        enunciado2 = New SqlCommand("select talm_15m as almacen, talm_15m + ' | '+ nomb_15m as descrip from custom_vianny.dbo.Mag1500 where ccia ='" + CCIA.Text + "' and  talm_15m ='" + ALMACEN.Text + "'", conx)
        respuesta2 = enunciado2.ExecuteReader
        While respuesta2.Read
            ComboBox1.Text = respuesta2.Item("descrip")
        End While
        respuesta2.Close()
        TextBox2.Select()
        If Label3.Text = 11 Or Label3.Text = 7 Then
            DataGridView1.Columns(0).Visible = False
            Button2.Enabled = False
        Else
            DataGridView1.Columns(0).Visible = True
            Button2.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = ""
        da.Clear()
        abrir()
        llenar_combo_box2()
        Dim cmd As New SqlDataAdapter("EXEC custom_vianny.dbo.CaeSoft_AyudaStockxAlmacen '" + CCIA.Text + "' , '" + TextBox1.Text + "', NULL, '" + ANO.Text + "', '" + Replace(FECHA.Text, "-", "") + "', NULL", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(2).Width = 180
            DataGridView1.Columns(4).Width = 450
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(11).Visible = True
            DataGridView1.Columns(13).Frozen = True
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            DataGridView1.Columns(4).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            DataGridView1.Columns(6).Frozen = True
            DataGridView1.Columns(7).Frozen = True
            DataGridView1.Columns(8).Frozen = True
            DataGridView1.Columns(9).Frozen = True
            DataGridView1.Columns(10).Frozen = True
            DataGridView1.Columns(11).Frozen = True
            DataGridView1.Columns(12).Frozen = True

        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Text = ComboBox1.SelectedValue.ToString
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If DataGridView1.Rows.Count > 0 Then
            bs.Filter = String.Format("Linea_17 LIKE '%{0}%'", TextBox3.Text)
        End If

    End Sub

    Sub llenar_combo_box2()

        Try

            conn = New SqlDataAdapter(" select talm_15m as almacen, talm_15m + ' | '+ nomb_15m as descrip from custom_vianny.dbo.Mag1500 where ccia ='" + CCIA.Text + "'", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "descrip"
            ComboBox1.ValueMember = "almacen"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Label3.Text = 11 Then
            OD.TextBox95.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString())
            OD.TextBox75.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value.ToString())
        Else
            If Label3.Text = "16" Then
                Padre.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString())
                Padre.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value.ToString())
            Else
                If Label3.Text = "12" Then
                    Matriz_Avios.DataGridView1.Rows(Label5.Text).Cells(6).Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value
                    Matriz_Avios.DataGridView1.Rows(Label5.Text).Cells(7).Value = DataGridView1.Rows(e.RowIndex).Cells(4).Value
                Else
                    If Label3.Text = "13" Then
                        DetalleMatrizAvios.DataGridView1.Rows(Label5.Text).Cells(3).Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value
                        DetalleMatrizAvios.DataGridView1.Rows(Label5.Text).Cells(4).Value = DataGridView1.Rows(e.RowIndex).Cells(4).Value
                    Else
                        If Label3.Text = "15" Then
                            ConsultarConsumo.TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString().Trim()
                        Else
                            If Label3.Text = "155" Then
                                Reporte_Ingreso_Salida.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString().Trim()
                                Reporte_Ingreso_Salida.TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value.ToString().Trim()
                            Else
                                Orden_Compra_Ac.TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString().Trim()
                            End If

                        End If

                    End If

                End If

            End If

        End If

        Me.Close()
    End Sub

    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit

    End Sub

    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            With DataGridView1

                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                End If
                If Label3.Text = 2 Then
                    Od_Udp.TextBox24.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(2).Value)
                    Od_Udp.TextBox25.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(4).Value)
                Else
                    OD.TextBox95.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(2).Value)
                    OD.TextBox75.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(4).Value)
                End If



            End With
            Me.Close()
        End If
    End Sub
End Class