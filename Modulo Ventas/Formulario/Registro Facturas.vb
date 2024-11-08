Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Net.Mail
Public Class Registro_Facturas

    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")


            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE rpm_ven ='2' AND ccia_ven ='" + Label29.Text + "'", conx)
            conn.Fill(ty2)
            ComboBox6.DataSource = ty2
            ComboBox6.DisplayMember = "alias_ven"
            ComboBox6.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Registro_Facturas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mostrar_corrlativo()
        ComboBox1.SelectedIndex = 0
        TextBox21.Select()
        abrir()
        TextBox14.Text = ""
        llenar_combo_box2()
        Dim func As New ftcambio
        Dim dts As New vtcambio
        dts.gfecha = DateTimePicker1.Text
        RadioButton6.Checked = True
        TextBox11.Text = func.mostrar_tipo_cambio(dts)
        If TextBox10.Text = "61" Then
            TextBox16.Text = "32"
            TextBox17.Text = "MUESTRAS SIN VALOR COMERCIAL - SOLES"
            RadioButton3.Checked = True
            RadioButton5.Visible = True
            RadioButton6.Visible = True
        Else
            RadioButton5.Visible = False
            RadioButton6.Visible = False
            TextBox16.Text = "04"
            TextBox17.Text = "VENTA LOCAL EN SOLES"
            RadioButton3.Checked = True
        End If
        Label28.Text = 1
    End Sub
    Private Sub mostrar_corrlativo()
        Try
            Dim func As New ffacturavianny
            Dim dts As New vfacturavianny

            TextBox2.Text = DateTime.Now.ToString("MM")
            dts.gmes = DateTime.Now.ToString("MM")
            dts.galmacen = Label5.Text
            dts.gano = Label27.Text
            dts.gccia = Label29.Text
            Me.TextBox1.Text = func.buscar_correlativo_venta(dts) + 1
            Select Case TextBox1.Text.Length

                Case "1" : TextBox1.Text = "00000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "0000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "6" : TextBox1.Text = TextBox1.Text
            End Select

        Catch ex As Exception



        End Try
    End Sub


    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Try
                Dim func As New ffacturavianny
                Dim dts As New vfacturavianny

                Select Case TextBox2.Text.Length

                    Case "1" : TextBox2.Text = "0" & "" & TextBox2.Text
                    Case "2" : TextBox2.Text = TextBox2.Text

                End Select
                dts.gmes = TextBox2.Text
                dts.galmacen = Label5.Text
                dts.gano = Label27.Text
                dts.gccia = Label29.Text
                Me.TextBox1.Text = func.buscar_correlativo_venta(dts) + 1
                Select Case TextBox1.Text.Length

                    Case "1" : TextBox1.Text = "00000" & "" & TextBox1.Text
                    Case "2" : TextBox1.Text = "0000" & "" & TextBox1.Text
                    Case "3" : TextBox1.Text = "000" & "" & TextBox1.Text
                    Case "4" : TextBox1.Text = "00" & "" & TextBox1.Text
                    Case "5" : TextBox1.Text = "0" & "" & TextBox1.Text
                    Case "6" : TextBox1.Text = TextBox1.Text
                End Select
                ComboBox1.Enabled = True
                TextBox21.ReadOnly = False
                TextBox21.Select()
            Catch ex As Exception



            End Try
        Else
            If e.KeyCode = Keys.F1 Then
                BUSCAR_VENTA.TextBox2.Text = Label17.Text
                BUSCAR_VENTA.TextBox3.Text = TextBox16.Text
                BUSCAR_VENTA.Label1.Text = 3
                BUSCAR_VENTA.Show()
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "FAC" Then
            TextBox31.Text = "001"
            TextBox30.Enabled = False
            TextBox33.Enabled = False
            ComboBox2.Enabled = False
        Else
            If ComboBox1.Text = "BOL" Then
                TextBox31.Text = "003"
                TextBox30.Enabled = False
                TextBox33.Enabled = False
                ComboBox2.Enabled = False
            Else
                If ComboBox1.Text = "NCR" Then
                    TextBox31.Text = "007"
                    TextBox30.Enabled = True
                    TextBox33.Enabled = True
                    ComboBox2.Enabled = True
                End If
            End If
        End If
    End Sub



    Private Sub TextBox21_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox21.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim func As New vfacturavianny
            Dim func1 As New ffacturavianny
            func.gserie = TextBox21.Text
            func.gdoc = TextBox31.Text

            func.gccia = Label29.Text
            TextBox5.Text = func1.correlativo_fac_bol(func) + 1

            Select Case TextBox5.Text.Length

                Case "1" : TextBox5.Text = "0000000" & "" & TextBox5.Text
                Case "2" : TextBox5.Text = "000000" & "" & TextBox5.Text
                Case "3" : TextBox5.Text = "00000" & "" & TextBox5.Text
                Case "4" : TextBox5.Text = "0000" & "" & TextBox5.Text
                Case "5" : TextBox5.Text = "000" & "" & TextBox5.Text
                Case "6" : TextBox5.Text = "00" & "" & TextBox5.Text
                Case "7" : TextBox5.Text = "0" & "" & TextBox5.Text
                Case "8" : TextBox5.Text = TextBox5.Text
            End Select
            If TextBox31.Text = "007" Then
                TextBox30.Select()
            Else
                TextBox18.ReadOnly = False
                TextBox19.ReadOnly = False
                TextBox18.Select()
            End If

        End If
    End Sub



    Private Sub TextBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox18.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox18.Text.Length
                Case "1" : TextBox18.Text = "000" & "" & TextBox18.Text
                Case "2" : TextBox18.Text = "00" & "" & TextBox18.Text
                Case "3" : TextBox18.Text = "0" & "" & TextBox18.Text
                Case "4" : TextBox18.Text = TextBox18.Text
            End Select
            TextBox19.Select()
        End If
    End Sub


    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox19.KeyDown

        If e.KeyCode = Keys.Enter Then
            'DataGridView1.Rows.Clear()
            Dim NU As Integer
            Dim hj As New fventasn
            Dim hj1 As New vventas_n
            Dim hj2 As New vventas_n

            Select Case TextBox19.Text.Length
                Case "1" : TextBox19.Text = "0000000" & "" & TextBox19.Text
                Case "2" : TextBox19.Text = "000000" & "" & TextBox19.Text
                Case "3" : TextBox19.Text = "00000" & "" & TextBox19.Text
                Case "4" : TextBox19.Text = "0000" & "" & TextBox19.Text
                Case "5" : TextBox19.Text = "000" & "" & TextBox19.Text
                Case "6" : TextBox19.Text = "00" & "" & TextBox19.Text
                Case "6" : TextBox19.Text = "0" & "" & TextBox19.Text
                Case "7" : TextBox19.Text = TextBox19.Text
            End Select
            hj1.gguia_nsa = TextBox18.Text & "-" & TextBox19.Text
            hj1.gccia = Label29.Text

            hj2.gccia = Label29.Text
            hj2.gguia_nsa = TextBox18.Text & "-" & TextBox19.Text
            Label31.Text = hj.buscartdoc(hj2)

            If Label31.Text = "01" And Microsoft.VisualBasic.Left(TextBox21.Text, 1) = "F" Then
                MsgBox("NO SE PERMITE REALIZAR FACTURAS A CLIENTES CON DNI, CAMBIE A BOLETA O INGRESE UN RUC")
            Else



                NU = hj.verificar_guia_vianny(hj1)

                If NU > 0 Then
                    Dim respuesta As Integer

                    respuesta = MsgBox("LA GUIA YA ESTA REGISTRADA, DESEA AGREGARLA?", vbQuestion + vbOKCancel)
                    If respuesta = 1 Then '-- 2=Cancelar;1 = Aceptar
                        '---- Acciones

                        cabecera_guia()
                        mostrar_detalle_guia()

                        TextBox13.ReadOnly = False
                        TextBox15.ReadOnly = False
                        TextBox16.ReadOnly = False
                        RadioButton1.Enabled = True
                        RadioButton2.Enabled = True
                        RadioButton3.Enabled = True
                        RadioButton4.Enabled = True
                        CheckBox1.Enabled = True
                        CheckBox2.Enabled = True
                        TextBox23.Enabled = True

                        DataGridView1.Enabled = True
                        DataGridView1.ReadOnly = False
                        DataGridView1.Columns(0).ReadOnly = True
                        DataGridView1.Columns(4).ReadOnly = False
                        DataGridView1.Columns(5).ReadOnly = False
                        RadioButton1.Checked = True
                        CheckBox1.Checked = True
                        CheckBox2.Checked = True
                        TextBox22.ReadOnly = False
                        TextBox18.Text = ""
                        TextBox19.Text = ""
                        TextBox18.Select()
                        Button4.Enabled = True
                        DataGridView2.DataSource = ""
                        DataGridView3.DataSource = ""
                    End If

                Else

                    cabecera_guia()
                    mostrar_detalle_guia()

                    TextBox13.ReadOnly = False
                    TextBox15.ReadOnly = False
                    TextBox16.ReadOnly = False
                    RadioButton1.Enabled = True
                    RadioButton2.Enabled = True
                    RadioButton3.Enabled = True
                    RadioButton4.Enabled = True
                    CheckBox1.Enabled = True
                    CheckBox2.Enabled = True
                    TextBox23.Enabled = True

                    DataGridView1.Enabled = True
                    DataGridView1.ReadOnly = False
                    DataGridView1.Columns(0).ReadOnly = True
                    DataGridView1.Columns(4).ReadOnly = False
                    DataGridView1.Columns(5).ReadOnly = False
                    RadioButton1.Checked = True
                    CheckBox1.Checked = True
                    CheckBox2.Checked = True
                    TextBox22.ReadOnly = False
                    TextBox18.Text = ""
                    TextBox19.Text = ""
                    TextBox18.Select()
                    Button4.Enabled = True
                    DataGridView2.DataSource = ""
                    DataGridView3.DataSource = ""
                End If

            End If
        End If


    End Sub
    Dim DT, DT1 As New DataTable
    Private Sub cabecera_guia()
        Try
            Dim func As New fguia2
            Dim dts As New vguia2
            Dim HG As New fcliente
            Dim HG1 As New vcliente
            Dim JH As String
            func.gsfactu = TextBox18.Text
            Select Case TextBox19.Text.Length

                Case "1" : func.gnfactu = "0000000" & "" & TextBox19.Text
                Case "2" : func.gnfactu = "000000" & "" & TextBox19.Text
                Case "3" : func.gnfactu = "00000" & "" & TextBox19.Text
                Case "4" : func.gnfactu = "0000" & "" & TextBox19.Text
                Case "5" : func.gnfactu = "000" & "" & TextBox19.Text
                Case "6" : func.gnfactu = "00" & "" & TextBox19.Text
                Case "7" : func.gnfactu = "0" & "" & TextBox19.Text
                Case "8" : func.gnfactu = TextBox19.Text
            End Select
            func.gccia = Label29.Text

            DT = dts.consultar_cabecera_guia(func)
            If DT.Rows.Count <> 0 Then
                DataGridView2.DataSource = DT
                If TextBox6.Text = "" Then
                    TextBox6.Text = DataGridView2.Rows(0).Cells(4).Value
                    TextBox7.Text = DataGridView2.Rows(0).Cells(5).Value
                    DateTimePicker1.Text = DataGridView2.Rows(0).Cells(0).Value
                    HG1.gruc = DataGridView2.Rows(0).Cells(4).Value.ToString.Trim
                    JH = HG.BUSCAR_VENDEDOR_CLIENTE(HG1)
                    TextBox14.Text = JH

                    abrir()
                    Dim sql102 As String = "SELECT  alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  codigo_ven ='" + JH + "'  AND ccia_ven ='" + Label29.Text + "'"
                    Dim cmd102 As New SqlCommand(sql102, conx)
                    Rsr3 = cmd102.ExecuteReader()
                    If Rsr3.Read() = True Then
                        ComboBox6.Text = Rsr3(0)

                    End If

                    Rsr3.Close()

                    If JH = "False" Then
                        ComboBox6.Enabled = True
                    End If
                    TextBox9.Text = "1"
                Else
                    If TextBox6.Text = DataGridView2.Rows(0).Cells(4).Value Then
                        TextBox6.Text = DataGridView2.Rows(0).Cells(4).Value
                        TextBox7.Text = DataGridView2.Rows(0).Cells(5).Value
                        DateTimePicker1.Text = DataGridView2.Rows(0).Cells(0).Value
                        TextBox9.Text = "1"
                    Else
                        TextBox9.Text = "2"
                        MsgBox("NO SE PUEDE AGREGRA GUIAS DE DIFERENTE CLIENTES")
                    End If
                End If



            End If
        Catch ex As Exception



        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        Dim i, a9, fila As Integer
        Dim cant10, sum10, cant9, sum9, cant8, sum8, cant7, sum7 As Double
        i = DataGridView1.Rows.Count
        fila = e.RowIndex

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cantidad" Then
            Try
                DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(6).Value) * Val(DataGridView1.Rows(fila).Cells(5).Value)
                DataGridView1.Rows(fila).Cells(7).Value = Val(DataGridView1.Rows(fila).Cells(9).Value) / 1.18
                DataGridView1.Rows(fila).Cells(8).Value = Val(DataGridView1.Rows(fila).Cells(9).Value) - Val(DataGridView1.Rows(fila).Cells(7).Value)
                Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
            Catch ex As Exception
                MsgBox("NO SELECCIONO UNA FILA CORRECTA")
            End Try

        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Precio Unitario" Then
            Try
                DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(6).Value) * Val(DataGridView1.Rows(fila).Cells(5).Value)
                DataGridView1.Rows(fila).Cells(7).Value = Val(DataGridView1.Rows(fila).Cells(9).Value) / 1.18
                DataGridView1.Rows(fila).Cells(8).Value = Val(DataGridView1.Rows(fila).Cells(9).Value) - Val(DataGridView1.Rows(fila).Cells(7).Value)
                Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
            Catch ex As Exception
                MsgBox("NO SELECCIONO UNA FILA CORRECTA")
            End Try

        End If
        For a9 = 0 To i - 1
            cant10 = Val(DataGridView1.Rows(a9).Cells(7).Value)
            sum10 = cant10 + Val(sum10)
            cant9 = Val(DataGridView1.Rows(a9).Cells(8).Value)
            sum9 = cant9 + Val(sum9)
            cant8 = Val(DataGridView1.Rows(a9).Cells(9).Value)
            sum8 = cant8 + Val(sum8)
            cant7 = Val(DataGridView1.Rows(a9).Cells(5).Value)
            sum7 = cant7 + Val(sum7)
            'cant6 = Val(DataGridView1.Rows(a9).Cells(8).Value)
            'sum6 = cant6 + Val(sum6)
        Next
        TextBox29.Text = sum8.ToString("N2")
        TextBox28.Text = sum9.ToString("N2")

        TextBox26.Text = sum10.ToString("N2")
        TextBox25.Text = sum7.ToString("N2")
        'TextBox23.Text = a9

    End Sub

    Private Sub mostrar_detalle_guia()
        Try
            'Dim a9
            'Dim cant9, sum9 As Double
            Dim func1 As New fguia2
            Dim dts1 As New vguia2


            func1.gsfactu = TextBox18.Text
            Select Case TextBox19.Text.Length

                Case "1" : func1.gnfactu = "0000000" & "" & TextBox19.Text
                Case "2" : func1.gnfactu = "000000" & "" & TextBox19.Text
                Case "3" : func1.gnfactu = "00000" & "" & TextBox19.Text
                Case "4" : func1.gnfactu = "0000" & "" & TextBox19.Text
                Case "5" : func1.gnfactu = "000" & "" & TextBox19.Text
                Case "6" : func1.gnfactu = "00" & "" & TextBox19.Text
                Case "7" : func1.gnfactu = "0" & "" & TextBox19.Text
                Case "8" : func1.gnfactu = TextBox19.Text
            End Select
            func1.gccia = Label29.Text
            DT1 = dts1.consultar_detalle_guia(func1)
            If DT1.Rows.Count <> 0 Then

                DataGridView3.DataSource = DT1

                If TextBox9.Text = 1 Then
                    Dim w1, w2, num1, w3 As Integer

                    w2 = DataGridView1.Rows.Count
                    w1 = DataGridView3.Rows.Count
                    num1 = 0

                    Dim Rsr1991 As SqlDataReader
                    Dim sql1011 As String = "select nombre,codigo_contabilidad from TABLA_RUBROS where codigo ='" + Trim(DataGridView3.Rows(0).Cells(9).Value) + "'"
                    Dim cmd1011 As New SqlCommand(sql1011, conx)
                    Rsr1991 = cmd1011.ExecuteReader()
                    Dim l As Integer
                    l = DataGridView1.Rows.Count
                    If Rsr1991.Read() Then

                        TextBox13.Text = Trim(Rsr1991(0))
                        Label4.Text = Trim(Rsr1991(1))
                    Else
                        TextBox13.Text = ""
                    End If
                    Rsr1991.Close()
                    DataGridView1.Rows.Add(w1)
                    For i6 = w2 To w1 + w2 - 1
                        If w2 > 0 Then

                            'MsgBox(w2)
                            w3 = w2 + 1
                            DataGridView1.Rows(i6).Cells(0).Value = w3

                            DataGridView1.Rows(i6).Cells(10).Value = DataGridView3.Rows(i6 - w2).Cells(1).Value
                            If TextBox10.Text = "61" Then
                                DataGridView1.Rows(i6).Cells(1).Value = "0108"
                            Else
                                DataGridView1.Rows(i6).Cells(1).Value = Label4.Text
                            End If

                            DataGridView1.Rows(i6).Cells(2).Value = DataGridView3.Rows(i6 - w2).Cells(2).Value
                            DataGridView1.Rows(i6).Cells(3).Value = DataGridView3.Rows(i6 - w2).Cells(3).Value
                            DataGridView1.Rows(i6).Cells(4).Value = DataGridView3.Rows(i6 - w2).Cells(7).Value
                            DataGridView1.Rows(i6).Cells(5).Value = DataGridView3.Rows(i6 - w2).Cells(6).Value

                            DataGridView1.Rows(i6).Cells(11).Value = "009"
                            DataGridView1.Rows(i6).Cells(12).Value = TextBox18.Text & "-" & TextBox19.Text
                            DataGridView1.Rows(i6).Cells(13).Value = DataGridView3.Rows(i6 - w2).Cells(8).Value
                            DataGridView1.Rows(i6).Cells(14).Value = DataGridView3.Rows(i6 - w2).Cells(9).Value
                        Else

                            num1 = num1 + 1 + w2

                            DataGridView1.Rows(i6).Cells(0).Value = num1
                            DataGridView1.Rows(i6).Cells(10).Value = DataGridView3.Rows(i6 - w2).Cells(1).Value
                            If TextBox10.Text = "61" Then
                                DataGridView1.Rows(i6).Cells(1).Value = ""
                            Else
                                DataGridView1.Rows(i6).Cells(1).Value = Label4.Text
                            End If

                            DataGridView1.Rows(i6).Cells(2).Value = DataGridView3.Rows(i6 - w2).Cells(2).Value
                            DataGridView1.Rows(i6).Cells(3).Value = DataGridView3.Rows(i6 - w2).Cells(3).Value
                            DataGridView1.Rows(i6).Cells(4).Value = DataGridView3.Rows(i6 - w2).Cells(7).Value
                            DataGridView1.Rows(i6).Cells(5).Value = DataGridView3.Rows(i6 - w2).Cells(6).Value

                            DataGridView1.Rows(i6).Cells(11).Value = "009"
                            DataGridView1.Rows(i6).Cells(12).Value = TextBox18.Text & "-" & TextBox19.Text
                            DataGridView1.Rows(i6).Cells(13).Value = DataGridView3.Rows(i6 - w2).Cells(8).Value
                            DataGridView1.Rows(i6).Cells(14).Value = DataGridView3.Rows(i6 - w2).Cells(9).Value

                        End If





                    Next



                End If






            End If
        Catch ex As Exception

            MsgBox("NO EXISTE ")

        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
        If CheckBox1.Checked = False Then

            Dim i8 As Integer
            i8 = DataGridView1.RowCount
            For i = 0 To i8 - 1
                CheckBox2.Enabled = False
                CheckBox2.Checked = False
                DataGridView1.Rows(i).Cells(9).Value = DataGridView1.Rows(i).Cells(5).Value * DataGridView1.Rows(i).Cells(6).Value
                DataGridView1.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(5).Value * DataGridView1.Rows(i).Cells(6).Value
                Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                DataGridView1.Rows(i).Cells(8).Value = "0.00"

            Next

            For a9 = 0 To i8 - 1
                cant10 = Val(DataGridView1.Rows(a9).Cells(7).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(a9).Cells(8).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(a9).Cells(9).Value)
                sum8 = cant8 + Val(sum8)

            Next
            TextBox29.Text = sum8.ToString("N2")
            TextBox28.Text = sum9.ToString("N2")

            TextBox26.Text = sum10.ToString("N2")
        Else
            If CheckBox1.Checked = True Then
                CheckBox2.Enabled = True
                Dim i8 As Integer
                i8 = DataGridView1.RowCount
                For i = 0 To i8 - 1

                    DataGridView1.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(5).Value * DataGridView1.Rows(i).Cells(6).Value

                    DataGridView1.Rows(i).Cells(9).Value = DataGridView1.Rows(i).Cells(9).Value * 1.18
                    DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(9).Value - DataGridView1.Rows(i).Cells(7).Value
                    Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
                Next

                For a9 = 0 To i8 - 1
                    cant10 = Val(DataGridView1.Rows(a9).Cells(7).Value)
                    sum10 = cant10 + Val(sum10)
                    cant9 = Val(DataGridView1.Rows(a9).Cells(8).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(DataGridView1.Rows(a9).Cells(9).Value)
                    sum8 = cant8 + Val(sum8)

                Next
                TextBox29.Text = sum8.ToString("N2")
                TextBox28.Text = sum9.ToString("N2")

                TextBox26.Text = sum10.ToString("N2")
            End If

        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
        If CheckBox2.Checked = True Then
            Dim i8 As Integer
            i8 = DataGridView1.RowCount
            For i = 0 To i8 - 1
                DataGridView1.Rows(i).Cells(9).Value = DataGridView1.Rows(i).Cells(5).Value * DataGridView1.Rows(i).Cells(6).Value
                DataGridView1.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(9).Value / 1.18
                DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(9).Value - DataGridView1.Rows(i).Cells(7).Value
                Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"

            Next

            For a9 = 0 To i8 - 1
                cant10 = Val(DataGridView1.Rows(a9).Cells(7).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(a9).Cells(8).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(a9).Cells(9).Value)
                sum8 = cant8 + Val(sum8)

            Next
            TextBox29.Text = sum8.ToString("N2")
            TextBox28.Text = sum9.ToString("N2")

            TextBox26.Text = sum10.ToString("N2")
        Else
            If CheckBox2.Checked = False And CheckBox1.Checked = True Then

                Dim i8 As Integer
                i8 = DataGridView1.RowCount
                For i = 0 To i8 - 1

                    DataGridView1.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(5).Value * DataGridView1.Rows(i).Cells(6).Value

                    DataGridView1.Rows(i).Cells(9).Value = DataGridView1.Rows(i).Cells(9).Value * 1.18
                    DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(9).Value - DataGridView1.Rows(i).Cells(7).Value
                    Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
                Next

                For a9 = 0 To i8 - 1
                    cant10 = Val(DataGridView1.Rows(a9).Cells(7).Value)
                    sum10 = cant10 + Val(sum10)
                    cant9 = Val(DataGridView1.Rows(a9).Cells(8).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(DataGridView1.Rows(a9).Cells(9).Value)
                    sum8 = cant8 + Val(sum8)

                Next
                TextBox29.Text = sum8.ToString("N2")
                TextBox28.Text = sum9.ToString("N2")

                TextBox26.Text = sum10.ToString("N2")
            End If
        End If
    End Sub






    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim respuesta As DialogResult
        Dim I18 As Integer

        Dim cant10, sum10, cant9, sum9, cant8, sum8, cant7, sum7 As Double

        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then

            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

            I18 = DataGridView1.Rows.Count
            For i1 = 0 To I18 - 1

                DataGridView1.Rows(i1).Cells(0).Value = i1 + 1
                cant10 = Val(DataGridView1.Rows(i1).Cells(7).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(i1).Cells(8).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(i1).Cells(9).Value)
                sum8 = cant8 + Val(sum8)
                cant7 = Val(DataGridView1.Rows(i1).Cells(5).Value)
                sum7 = cant7 + Val(sum7)
            Next
            TextBox29.Text = sum8.ToString("N2")
            TextBox28.Text = sum9.ToString("N2")
            TextBox26.Text = sum10.ToString("N2")
            TextBox25.Text = sum7.ToString("N2")
        End If

    End Sub
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        'Dim corre As String
        'jj.gvendedor = TextBox8.Text
        'corre = fk.buscar_correo(jj)
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("hrivera@viannysac.com,asistentecobranzas@viannysac.com")
            .Body = "Nota Genero el Comprobante N°" & ComboBox1.Text & "-" & TextBox21.Text & "-" & TextBox5.Text
            .Subject = "Cliente" & TextBox7.Text & "-   Almacen" & TextBox10.Text & "-" & "F_Pagp" & TextBox23.Text
            .Priority = System.Net.Mail.MailPriority.Normal
        End With

        With smtp
            .EnableSsl = True
            .Port = "25"
            .Host = "mail.onehostingperu.com"
            .Credentials = New Net.NetworkCredential("admin@viannysac.com", "I9!?@ni2;+go")
            .Send(message)
        End With

        Try
            MessageBox.Show("Su mensaje de correo ha sido enviado.", "Correo enviado",
                             MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error al enviar correo",
                             MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        abrir()
        Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE  alias_ven ='" + ComboBox6.Text + "' AND ccia_ven ='" + Label29.Text + "' "
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr4 = cmd102.ExecuteReader()
        If Rsr4.Read() = True Then
            TextBox14.Text = Rsr4(0)
        End If
        Rsr4.Close()
        'Select Case ComboBox6.Text
        '    Case "GBEDON" : TextBox14.Text = "0010"
        '    Case "VINCIO" : TextBox14.Text = "0022"
        '    Case "DBRAVO" : TextBox14.Text = "0023"
        '    Case "JSALINAS" : TextBox14.Text = "0025"
        '    Case "GCUEVA" : TextBox14.Text = "0029"
        '    Case "AMENDO" : TextBox14.Text = "0026"
        '    Case "VGRAUS" : TextBox14.Text = "0007"
        '    Case "VPIZARRO" : TextBox14.Text = "0027"
        '    Case "JBALVIN" : TextBox14.Text = "0028"
        '    Case "VSILVERIO" : TextBox14.Text = "0005"
        '    Case "RMEDINA" : TextBox14.Text = "0033"
        '    Case "WSALINAS" : TextBox14.Text = "0034"
        '    Case "MGRAUS" : TextBox14.Text = "0037"
        'End Select
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form_Cliente.Label1.Text = "F. PAGO"
        Form_Cliente.TextBox3.Text = 4

        Form_Cliente.Show()
    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)
        MsgBox("LA OPCION ESTA DESHABILITADA, SOLICITAR AL AREA CONTABLE LA ANULACION")
        'Dim respuesta As String
        'respuesta = MessageBox.Show("DESEA ANULAR LA FACTURA", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        'If (respuesta = Windows.Forms.DialogResult.Yes) Then

        '    Dim JH As New vfacturasistema
        '    Dim HL As New ffacturavianny
        '    JH.gflag = 0
        '    JH.gncom = TextBox2.Text & "" & TextBox1.Text
        '    JH.galmacen = TextBox10.Text
        '    JH.gano = Label27.Text
        '    JH.gccia = Label29.Text
        '    HL.anular_factura(JH)
        '    Label28.Text = 3

        '    Dim hjt As New fcontable
        '    Dim ggg As New vcontable
        '    Dim ju4 As New vcontable

        '    ju4.gALMACEN = TextBox10.Text
        '    ju4.gccia_3 = Label29.Text
        '    abrir()
        '    Dim agregar As String = "DELETE FROM custom_vianny.dbo.Cac3p00 WHERE ccia_3A = '" + Label29.Text + "' AND CPERIODO_3A = '" + Label27.Text + "' AND ccom_3a = '" + hjt.BUSCAR_SUBDIARIO(ju4) + "' AND NCOM_3A = '" + TextBox2.Text & "" & TextBox1.Text + "'"
        '    Dim agregar1 As String = "DELETE FROM custom_vianny.dbo.Cac3g00 WHERE ccia_3 = '" + Label29.Text + "' AND cperiodo_3 = '" + Label27.Text + "' AND ccom_3 = '" + hjt.BUSCAR_SUBDIARIO(ju4) + "' AND ncom_3 = '" + TextBox2.Text & "" & TextBox1.Text + "'"
        '    Dim agregar2 As String = "DELETE FROM custom_vianny.dbo.cac3r00 WHERE ccia_3a = '" + Label29.Text + "' AND cperiod_3a = '" + Label27.Text + "' AND ccom_3a = '" + hjt.BUSCAR_SUBDIARIO(ju4) + "' AND ncom_3a = '" + TextBox2.Text & "" & TextBox1.Text + "'"

        '    ELIMINAR(agregar)
        '    ELIMINAR(agregar1)
        '    ELIMINAR(agregar2)
        '    MsgBox("SE ANULO LA INFORMACION SOLICITADA")
        '    mostrar_corrlativo()
        '    TextBox21.Text = ""
        '    TextBox5.Text = ""
        '    TextBox6.Text = ""
        '    TextBox7.Text = ""
        '    TextBox23.Text = ""
        '    TextBox13.Text = ""
        '    TextBox15.Text = ""
        '    TextBox22.Text = ""
        '    TextBox25.Text = ""
        '    TextBox26.Text = ""
        '    TextBox28.Text = ""
        '    TextBox29.Text = ""
        '    TextBox14.Text = ""
        '    TextBox27.Text = ""
        '    TextBox30.Text = ""
        '    TextBox33.Text = ""
        '    TextBox34.Text = ""
        '    ComboBox2.Text = ""
        '    DataGridView1.Rows.Clear()
        '    TextBox21.Enabled = False
        '    TextBox5.Enabled = False
        '    TextBox6.Enabled = False
        '    TextBox7.Enabled = False
        '    TextBox23.Enabled = False
        '    TextBox13.Enabled = False
        '    TextBox15.Enabled = False
        '    TextBox22.Enabled = False
        '    TextBox25.Enabled = False
        '    TextBox26.Enabled = False
        '    TextBox28.Enabled = False
        '    TextBox29.Enabled = False
        '    TextBox14.Enabled = False
        '    ComboBox1.SelectedIndex = 0
        '    DataGridView1.Enabled = False
        'End If
    End Sub
    Public comando As SqlCommand

    Function ELIMINAR(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If TextBox10.Text = "61" Then
            If RadioButton5.Checked = True Then
                TextBox16.Text = "03"
                TextBox17.Text = "VENTAS LOCALES - DOLARES"
            Else
                TextBox16.Text = "33"
                TextBox17.Text = "MUESTRAS SIN VALOR COMERCIAL - DOLARES"
            End If

        Else
            TextBox16.Text = "03"
            TextBox17.Text = "VENTAS LOCALES - DOLARES"
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If TextBox10.Text = "61" Then
            If RadioButton5.Checked = True Then
                TextBox16.Text = "04"
                TextBox17.Text = "VENTAS LOCALES - SOLES"
            Else
                TextBox16.Text = "32"
                TextBox17.Text = "MUESTRAS SIN VALOR COMERCIAL - SOLES"
            End If


        Else
            TextBox16.Text = "04"
            TextBox17.Text = "VENTAS LOCALES - SOLES"
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "FAC" Then
            TextBox27.Text = "001"
        Else
            If ComboBox2.Text = "BOL" Then
                TextBox27.Text = "003"
            End If
        End If
    End Sub
    Dim Rsr21 As SqlDataReader
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged

    End Sub

    Private Sub TextBox30_TextChanged(sender As Object, e As EventArgs) Handles TextBox30.TextChanged

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, a9, fila As Integer
        Dim cant10, sum10, cant9, sum9, cant8, sum8, sum7, cant7 As Double
        i = DataGridView1.Rows.Count
        fila = e.RowIndex

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cantidad" Then
            DataGridView1.Rows(fila).Cells(9).Value = DataGridView1.Rows(fila).Cells(6).Value * DataGridView1.Rows(fila).Cells(5).Value
            DataGridView1.Rows(fila).Cells(7).Value = DataGridView1.Rows(fila).Cells(9).Value / 1.18
            DataGridView1.Rows(fila).Cells(8).Value = DataGridView1.Rows(fila).Cells(9).Value - DataGridView1.Rows(fila).Cells(7).Value
            Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
            Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n2"
            Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Precio Unitario" Then
            DataGridView1.Rows(fila).Cells(9).Value = DataGridView1.Rows(fila).Cells(6).Value * DataGridView1.Rows(fila).Cells(5).Value
            DataGridView1.Rows(fila).Cells(7).Value = DataGridView1.Rows(fila).Cells(9).Value / 1.18
            DataGridView1.Rows(fila).Cells(8).Value = DataGridView1.Rows(fila).Cells(9).Value - DataGridView1.Rows(fila).Cells(7).Value
            Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
            Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
            Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
        End If
        For a9 = 0 To i - 1
            cant10 = Val(DataGridView1.Rows(a9).Cells(7).Value)
            sum10 = cant10 + Val(sum10)
            cant9 = Val(DataGridView1.Rows(a9).Cells(8).Value)
            sum9 = cant9 + Val(sum9)
            cant8 = Val(DataGridView1.Rows(a9).Cells(9).Value)
            sum8 = cant8 + Val(sum8)
            cant7 = Val(DataGridView1.Rows(a9).Cells(5).Value)
            sum7 = cant7 + Val(sum7)

        Next
        TextBox29.Text = sum8.ToString("N2")
        TextBox28.Text = sum9.ToString("N2")

        TextBox26.Text = sum10.ToString("N2")
        TextBox25.Text = sum7.ToString("N2")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub TextBox33_TextChanged(sender As Object, e As EventArgs) Handles TextBox33.TextChanged

    End Sub

    Dim gg, gg1 As New DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim jj As New ffacturavianny
            Dim ll As New vfacturasistema
            Dim y, l As Integer
            Select Case TextBox1.Text.Length

                Case "1" : TextBox1.Text = "00000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "0000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "6" : TextBox1.Text = TextBox1.Text
            End Select
            ll.gccia = Label29.Text
            ll.gncom = TextBox2.Text & "" & TextBox1.Text
            ll.galmacen = TextBox10.Text
            ll.gano = Label27.Text
            DataGridView1.Rows.Clear()
            gg = jj.mostrar_factura_vianny_cabecera(ll)

            gg1 = jj.mostrar_factura_vianny_detalle(ll)

            DataGridView4.DataSource = gg
            DataGridView5.DataSource = gg1
            l = DataGridView4.Rows.Count
            If l > 0 Then
                y = DataGridView5.Rows.Count
                DataGridView1.Rows.Add(y)

                TextBox31.Text = Trim(DataGridView4.Rows(0).Cells(0).Value)
                Select Case TextBox31.Text

                    Case "001" : ComboBox1.Text = "FAC"
                    Case "003" : ComboBox1.Text = "BOL"
                    Case "007" : ComboBox1.Text = "NCR"

                End Select
                TextBox34.Text = Trim(DataGridView4.Rows(0).Cells(20).Value)
                TextBox21.Text = DataGridView4.Rows(0).Cells(1).Value
                TextBox5.Text = DataGridView4.Rows(0).Cells(2).Value
                TextBox6.Text = DataGridView4.Rows(0).Cells(3).Value
                TextBox7.Text = DataGridView4.Rows(0).Cells(4).Value
                DateTimePicker1.Value = DataGridView4.Rows(0).Cells(5).Value
                If DataGridView4.Rows(0).Cells(6).Value Is DBNull.Value Then

                Else
                    DateTimePicker2.Value = DataGridView4.Rows(0).Cells(6).Value
                End If

                TextBox11.Text = DataGridView4.Rows(0).Cells(16).Value
                TextBox20.Text = DataGridView4.Rows(0).Cells(7).Value
                abrir()
                enunciado2 = New SqlCommand("SELECT  nomb_15k  FROM custom_vianny.dbo.kag1500 A  Where A.ccia_15k='01' and flag_15k='1' and cond_15k ='" + DataGridView4.Rows(0).Cells(7).Value + "'", conx)

                respuesta2 = enunciado2.ExecuteReader
                While respuesta2.Read
                    TextBox23.Text = respuesta2.Item("nomb_15k")
                End While
                respuesta2.Close()
                TextBox16.Text = DataGridView4.Rows(0).Cells(11).Value
                TextBox13.Text = DataGridView4.Rows(0).Cells(8).Value
                If DataGridView4.Rows(0).Cells(9).Value = 1 Then
                    CheckBox1.Checked = True
                Else
                    CheckBox1.Checked = False
                End If
                If DataGridView4.Rows(0).Cells(10).Value = 1 Then
                    CheckBox2.Checked = True
                Else
                    CheckBox2.Checked = False
                End If
                If DataGridView4.Rows(0).Cells(14).Value = 1 Then
                    RadioButton3.Checked = True
                    Label32.Visible = False
                Else
                    RadioButton4.Checked = True
                    Label32.Visible = True
                End If
                If DataGridView4.Rows(0).Cells(15).Value = 1 Then
                    RadioButton1.Checked = True
                Else
                    RadioButton2.Checked = True
                End If
                TextBox22.Text = DataGridView4.Rows(0).Cells(13).Value
                TextBox26.Text = DataGridView4.Rows(0).Cells(17).Value
                TextBox28.Text = DataGridView4.Rows(0).Cells(18).Value
                TextBox29.Text = DataGridView4.Rows(0).Cells(19).Value
                TextBox14.Text = DataGridView4.Rows(0).Cells(12).Value

                abrir()
                enunciado3 = New SqlCommand("select alias_ven from custom_vianny.dbo.Vendedores where codigo_ven = " + DataGridView4.Rows(0).Cells(12).Value, conx)

                respuesta3 = enunciado3.ExecuteReader

                While respuesta3.Read

                    'ComboBox6.Items.Clear()
                    'ComboBox6.Items.Add(respuesta3.Item("alias_ven"))
                    'ComboBox6.SelectedIndex = 0
                    ComboBox6.Text = respuesta3.Item("alias_ven")
                End While
                respuesta3.Close()
                For i = 0 To y - 1
                    DataGridView1.Rows(i).Cells(0).Value = DataGridView5.Rows(i).Cells(0).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView5.Rows(i).Cells(1).Value
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView5.Rows(i).Cells(2).Value
                    DataGridView1.Rows(i).Cells(3).Value = DataGridView5.Rows(i).Cells(3).Value
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView5.Rows(i).Cells(4).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView5.Rows(i).Cells(5).Value
                    DataGridView1.Rows(i).Cells(6).Value = DataGridView5.Rows(i).Cells(6).Value
                    DataGridView1.Rows(i).Cells(7).Value = DataGridView5.Rows(i).Cells(7).Value
                    DataGridView1.Rows(i).Cells(8).Value = DataGridView5.Rows(i).Cells(8).Value
                    DataGridView1.Rows(i).Cells(9).Value = DataGridView5.Rows(i).Cells(9).Value
                    DataGridView1.Rows(i).Cells(10).Value = DataGridView5.Rows(i).Cells(10).Value
                    DataGridView1.Rows(i).Cells(11).Value = DataGridView5.Rows(i).Cells(11).Value
                    DataGridView1.Rows(i).Cells(12).Value = DataGridView5.Rows(i).Cells(12).Value
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView5.Rows(i).Cells(13).Value
                    DataGridView1.Rows(i).Cells(14).Value = DataGridView5.Rows(i).Cells(14).Value
                Next
                Button7.Visible = True
                Button9.Visible = True
                Button5.Enabled = False
                Button8.Enabled = False
                Button3.Enabled = False
                Button6.Enabled = False
                TextBox21.Enabled = False
                TextBox5.Enabled = False
                DateTimePicker1.Enabled = False
                DateTimePicker2.Enabled = False
                ComboBox1.Enabled = False
                CheckBox2.Enabled = False
            Else
                Button6.Enabled = True
                TextBox21.Select()
                TextBox21.Enabled = True
                TextBox5.Enabled = True
                TextBox21.Text = ""
                TextBox5.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = ""
                TextBox13.Text = ""
                TextBox22.Text = ""
                TextBox14.Text = ""
                TextBox20.Text = ""
                TextBox26.Text = ""
                TextBox23.Text = ""
                TextBox28.Text = ""
                TextBox29.Text = ""
            End If

        Else
            If e.KeyCode = Keys.F1 Then
                TextBox21.Text = ""
                TextBox5.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = ""
                TextBox23.Text = ""
                TextBox13.Text = ""
                TextBox15.Text = ""
                TextBox22.Text = ""
                TextBox25.Text = ""
                TextBox26.Text = ""
                TextBox28.Text = ""
                TextBox29.Text = ""
                TextBox14.Text = ""
                TextBox27.Text = ""
                TextBox30.Text = ""
                TextBox33.Text = ""
                TextBox34.Text = ""
                ComboBox2.Text = ""
                DataGridView1.Rows.Clear()
                BUSCAR_VENTA.TextBox2.Text = Label17.Text
                BUSCAR_VENTA.TextBox3.Text = TextBox16.Text
                BUSCAR_VENTA.Label1.Text = 2
                BUSCAR_VENTA.Show()
            End If


        End If

    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim func As New ftcambio
        Dim dts As New vtcambio
        dts.gfecha = DateTimePicker1.Text
        TextBox11.Text = func.mostrar_tipo_cambio(dts)

    End Sub

    Private Sub TextBox30_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox30.KeyDown
        If e.KeyCode = Keys.Enter Then
            'Dim func As New vfacturavianny
            'Dim func1 As New ffacturavianny
            'func.gserie = TextBox30.Text
            'func.gdoc = TextBox27.Text
            'TextBox33.Text = func1.correlativo_fac_bol(func) + 1

            'Select Case TextBox33.Text.Length

            '    Case "1" : TextBox33.Text = "0000000" & "" & TextBox33.Text
            '    Case "2" : TextBox33.Text = "000000" & "" & TextBox33.Text
            '    Case "3" : TextBox33.Text = "00000" & "" & TextBox33.Text
            '    Case "4" : TextBox33.Text = "0000" & "" & TextBox33.Text
            '    Case "5" : TextBox33.Text = "000" & "" & TextBox33.Text
            '    Case "6" : TextBox33.Text = "00" & "" & TextBox33.Text
            '    Case "7" : TextBox33.Text = "0" & "" & TextBox33.Text
            '    Case "8" : TextBox33.Text = TextBox33.Text
            'End Select

            'TextBox18.ReadOnly = False
            'TextBox19.ReadOnly = False
            TextBox33.Select()
        End If
    End Sub
    Dim Rsr2, Rsr3, Rsr4, Rsr5 As SqlDataReader
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim cab As New vfacturasistema
        Dim cab1 As New vfacturasistema
        Dim det As New vfacturadetallesistema
        Dim cla As New ffacturavianny
        Dim ad, MJ, ca As Integer

        Dim suma, suma1, suma2, suma3, suma4, suma5, tcambio As Double
        ca = 0
        MJ = DataGridView1.Rows.Count
        For i1 = 0 To MJ - 1
            If Mid(Trim(DataGridView1.Rows(i1).Cells(1).Value), 1, 4) = "Labe" Or Mid(Trim(DataGridView1.Rows(i1).Cells(1).Value), 1, 4) = "LABE" Then
                ca = ca + 1
            End If
        Next

        If ca >= 1 Then
            MsgBox("EL RUBRO NO ESTA INGRESADO CORRECTAMENTE, INFORME AL AREA DE SISTEMAS PARA SU CORRECCION")
        Else




            If Trim(TextBox34.Text) = "" Then
                MsgBox("ES OBLIGATORIO INGRESAR LA NOTA DE PEDIDO")
            Else


                If Trim(TextBox23.Text) = "" Or Trim(TextBox6.Text) = "" Or Trim(TextBox5.Text) = "" Or Trim(ComboBox6.Text) = "" Or MJ = 0 Then
                    MsgBox("ALGUNOS CAMPOS OBLIGATORIOS ESTAN VACIOS, EL TIPO DE CAMBIO TIENE QUE SER MAYOR A 0 O FALTA INGRESAR EL DETALLE EN LA FACTURA")

                Else


                    tcambio = TextBox11.Text
                    If Label28.Text = 1 Then
                        Dim func As New ffacturavianny
                        Dim dts As New vfacturavianny
                        Dim np As Int16
                        Select Case TextBox2.Text.Length

                            Case "1" : TextBox2.Text = "0" & "" & TextBox2.Text
                            Case "2" : TextBox2.Text = TextBox2.Text

                        End Select
                        dts.gmes = TextBox2.Text
                        dts.galmacen = Label5.Text
                        dts.gano = Label27.Text
                        dts.gccia = Label29.Text
                        np = func.buscar_correlativo_venta(dts) + 1
                        Select Case np.ToString.Length

                            Case "1" : cab.gncom = TextBox2.Text & "" & "00000" & "" & np
                            Case "2" : cab.gncom = TextBox2.Text & "" & "0000" & "" & np
                            Case "3" : cab.gncom = TextBox2.Text & "" & "000" & "" & np
                            Case "4" : cab.gncom = TextBox2.Text & "" & "00" & "" & np
                            Case "5" : cab.gncom = TextBox2.Text & "" & "0" & "" & np
                            Case "6" : cab.gncom = TextBox2.Text & "" & np
                        End Select

                    Else
                        cab.gncom = TextBox2.Text & "" & TextBox1.Text
                    End If

                    cab.gfatura = TextBox21.Text
                    cab.gcorrelativo_fac = TextBox5.Text
                    cab.gfecha = DateTimePicker1.Value
                    cab.gcondicion_pago = TextBox20.Text
                    cab.gruc = Trim(TextBox6.Text)
                    cab.grazon_social = TextBox7.Text

                    If RadioButton1.Checked = True Then
                        cab.gmoneda = 1
                    Else
                        If RadioButton2.Checked = True Then
                            cab.gmoneda = 2
                        End If
                    End If
                    cab.gcambio = TextBox11.Text
                    '-------------------------------------

                    ad = DataGridView1.Rows.Count
                    If RadioButton1.Checked = True Then
                        For i = 0 To ad - 1

                            suma = suma + Val(DataGridView1.Rows(i).Cells(7).Value)
                            suma1 = suma1 + Val(DataGridView1.Rows(i).Cells(8).Value)
                            suma2 = suma2 + Val(DataGridView1.Rows(i).Cells(9).Value)
                        Next
                        suma3 = suma / tcambio
                        suma4 = suma1 / tcambio
                        suma5 = suma2 / tcambio

                        cab.gventa_total = suma2
                        cab.gventa_sinigv = suma
                        cab.gigv = suma1
                        cab.gtotal_venta = suma2

                        cab.gventa_sinigv_do = suma3
                        cab.gigv_do = suma4
                        cab.gtotal_venta_do = suma5
                    Else
                        If RadioButton2.Checked = True Then
                            For i = 0 To ad - 1

                                suma = suma + Val(DataGridView1.Rows(i).Cells(7).Value)
                                suma1 = suma1 + Val(DataGridView1.Rows(i).Cells(8).Value)
                                suma2 = suma2 + Val(DataGridView1.Rows(i).Cells(9).Value)
                            Next
                            suma3 = suma * tcambio
                            suma4 = suma1 * tcambio
                            suma5 = suma2 * tcambio
                            cab.gventa_total = suma5
                            cab.gventa_sinigv = suma3
                            cab.gigv = suma4
                            cab.gtotal_venta = suma5

                            cab.gventa_sinigv_do = suma
                            cab.gigv_do = suma1
                            cab.gtotal_venta_do = suma2
                        End If
                    End If

                    '------------------------------------


                    cab.gglosa = Trim(Microsoft.VisualBasic.Mid(TextBox13.Text, 1, 60))
                    cab.gvendedor = TextBox14.Text
                    cab.galmacen = TextBox10.Text
                    cab.gfechafin = DateTimePicker2.Value
                    If TextBox27.Text = "" Then
                        cab.gstdoc = ""
                    Else
                        cab.gstdoc = TextBox27.Text
                    End If
                    If TextBox30.Text = "" Then
                        cab.gsfatura = ""
                    Else
                        cab.gsfatura = TextBox30.Text
                    End If
                    If TextBox33.Text = "" Then
                        cab.gscorrelativo_fac = ""
                    Else
                        cab.gscorrelativo_fac = TextBox33.Text
                    End If
                    If TextBox31.Text = "007" Then
                        cab.goper = "05"
                    Else
                        cab.goper = ""
                    End If


                    If CheckBox1.Checked = True Then
                        cab.gafecto = 1
                    Else
                        cab.gafecto = 0
                    End If
                    If CheckBox2.Checked = True Then
                        cab.gincluido = 1
                    Else
                        cab.gincluido = 0
                    End If
                    cab.gpedido = TextBox34.Text
                    cab.gglosafac = TextBox22.Text
                    cab.gtipo_venta = TextBox16.Text
                    cab.gtdoc = TextBox31.Text
                    cab.gano = Label27.Text
                    cab.gccia = Label29.Text
                    Dim hh As New ffacturavianny
                    Dim hh1 As New vfacturasistema
                    hh1.gncom = TextBox2.Text & "" & TextBox1.Text
                    hh1.galmacen = TextBox10.Text
                    'hh.eliminar_regventanegr(hh1)
                    cab1.gncom = cab.gncom
                    cab1.galmacen = TextBox10.Text
                    cab1.gano = Label27.Text
                    cab1.gccia = Label29.Text
                    cla.eiminar_factura_vianny(cab1)
                    'eliminar registro contable

                    'INSERTAR REGISTRO CONTABLE
                    Dim hjt As New fcontable
                    Dim ggg As New vcontable
                    Dim ju4 As New vcontable
                    ggg.gcperiodo_3 = Label27.Text
                    ju4.gALMACEN = TextBox10.Text
                    ju4.gccia_3 = Label29.Text
                    ggg.gccom_3 = hjt.BUSCAR_SUBDIARIO(ju4)
                    ggg.gncom_3 = TextBox2.Text & "" & TextBox1.Text
                    ggg.gccia_3 = Label29.Text
                    hjt.eliminar_registro_contable(ggg)

                    '------------------
                    Dim lp As New vcontable
                    Dim lp1 As New fcontable
                    Dim ju As New vcontable
                    Dim YU As Integer
                    Dim sumo2, sumo5, tcambio1 As Double
                    tcambio1 = TextBox11.Text
                    lp.gccia_3 = Label29.Text

                    lp.gcperiodo_3 = Label27.Text

                    ju.gALMACEN = TextBox10.Text
                    ju.gccia_3 = Label29.Text
                    lp.gccom_3 = lp1.BUSCAR_SUBDIARIO(ju)

                    'If Label28.Text = 1 Then
                    '    Dim func As New ffacturavianny
                    '    Dim dts As New vfacturavianny
                    '    Dim np As Int16
                    '    Select Case TextBox2.Text.Length

                    '        Case "1" : TextBox2.Text = "0" & "" & TextBox2.Text
                    '        Case "2" : TextBox2.Text = TextBox2.Text

                    '    End Select
                    '    dts.gmes = TextBox2.Text
                    '    dts.galmacen = Label5.Text
                    '    dts.gano = Label27.Text
                    '    dts.gccia = Label29.Text
                    '    np = func.buscar_correlativo_venta(dts) + 1
                    '    Select Case np.ToString.Length

                    '        Case "1" : lp.gncom_3 = TextBox2.Text & "" & "00000" & "" & np
                    '        Case "2" : lp.gncom_3 = TextBox2.Text & "" & "0000" & "" & np
                    '        Case "3" : lp.gncom_3 = TextBox2.Text & "" & "000" & "" & np
                    '        Case "4" : lp.gncom_3 = TextBox2.Text & "" & "00" & "" & np
                    '        Case "5" : lp.gncom_3 = TextBox2.Text & "" & "0" & "" & np
                    '        Case "6" : lp.gncom_3 = TextBox2.Text & "" & np
                    '    End Select

                    'Else
                    '    lp.gncom_3 = TextBox2.Text & "" & TextBox1.Text
                    'End If
                    lp.gncom_3 = cab.gncom

                    lp.gtcom_3 = "2"

                    If RadioButton1.Checked = True Then
                        lp.gmone_3 = 1
                    Else
                        If RadioButton2.Checked = True Then
                            lp.gmone_3 = 2
                        End If
                    End If

                    lp.gfcom_3 = DateTimePicker1.Value

                    lp.gtcam_3 = TextBox11.Text

                    lp.gglosa_3 = Trim(Microsoft.VisualBasic.Mid(TextBox13.Text, 1, 60))

                    lp.gglosb_3 = ""

                    YU = DataGridView1.Rows.Count
                    If RadioButton1.Checked = True Then
                        For i = 0 To YU - 1

                            sumo2 = sumo2 + Val(DataGridView1.Rows(i).Cells(9).Value)
                        Next

                        sumo5 = sumo2 / tcambio1


                        lp.gtotd1_3 = sumo2
                        lp.gtotd2_3 = sumo5
                        lp.gtoth1_3 = sumo2
                        lp.gtoth2_3 = sumo5
                    Else
                        If RadioButton2.Checked = True Then
                            For i = 0 To YU - 1


                                sumo2 = sumo2 + Val(DataGridView1.Rows(i).Cells(9).Value)
                            Next

                            sumo5 = suma2 * tcambio1

                            lp.gtotd1_3 = sumo5
                            lp.gtotd2_3 = sumo2
                            lp.gtoth1_3 = sumo5
                            lp.gtoth2_3 = sumo2
                        End If
                    End If


                    lp.gcuenb_3 = ""

                    lp.ggirado_3 = ""

                    lp.gnombcp_3 = TextBox7.Text

                    lp.gfcoma_3 = DateTimePicker1.Value

                    lp.gnroa_3 = "001"

                    lp.gtmov_3 = ""

                    lp.gdifcam_3 = 0

                    lp.gtasien_3 = 3

                    lp.gflag_3 = 1

                    lp.gfcome_3 = DateTimePicker1.Value

                    lp.gusuario_3 = TextBox12.Text

                    lp.gfupdate_3 = ""

                    lp1.insertar_cabecera_contable(lp)

                    'fin regitro cabecer contabel

                    If cla.insertar_cabecera_venta(cab) = True Then
                        For i1 = 0 To ad - 1
                            Dim nu As String

                            'If Label28.Text = 1 Then
                            '    Dim func As New ffacturavianny
                            '    Dim dts As New vfacturavianny
                            '    Dim np As Int16
                            '    Select Case TextBox2.Text.Length

                            '        Case "1" : TextBox2.Text = "0" & "" & TextBox2.Text
                            '        Case "2" : TextBox2.Text = TextBox2.Text

                            '    End Select
                            '    dts.gmes = TextBox2.Text
                            '    dts.galmacen = Label5.Text
                            '    dts.gano = Label27.Text
                            '    dts.gccia = Label29.Text
                            '    np = func.buscar_correlativo_venta(dts) + 1
                            '    Select Case np.ToString.Length

                            '        Case "1" : det.gncom = TextBox2.Text & "" & "00000" & "" & np
                            '        Case "2" : det.gncom = TextBox2.Text & "" & "0000" & "" & np
                            '        Case "3" : det.gncom = TextBox2.Text & "" & "000" & "" & np
                            '        Case "4" : det.gncom = TextBox2.Text & "" & "00" & "" & np
                            '        Case "5" : det.gncom = TextBox2.Text & "" & "0" & "" & np
                            '        Case "6" : det.gncom = TextBox2.Text & "" & np
                            '    End Select

                            'Else
                            '    det.gncom = TextBox2.Text & "" & TextBox1.Text
                            'End If
                            det.gncom = cab.gncom
                            'det.gitems = DataGridView1.Rows(i1).Cells(0).Value
                            nu = DataGridView1.Rows(i1).Cells(0).Value

                            Select Case nu.Length
                                Case "1" : nu = "00" & "" & nu
                                Case "2" : nu = "0" & "" & nu
                                Case "3" : nu = nu
                            End Select
                            det.gitems = nu

                            det.glinea = DataGridView1.Rows(i1).Cells(2).Value
                            det.gsinonimo = "001"
                            det.gcantidad = DataGridView1.Rows(i1).Cells(5).Value

                            '--------
                            If RadioButton1.Checked = True Then
                                det.gprecio_unitario = Convert.ToDouble(DataGridView1.Rows(i1).Cells(6).Value)
                                det.gventa_total = (DataGridView1.Rows(i1).Cells(8).Value)
                                det.gvalor_venta = (DataGridView1.Rows(i1).Cells(7).Value)
                                det.gigv = (DataGridView1.Rows(i1).Cells(8).Value)
                                det.gtotal = (DataGridView1.Rows(i1).Cells(9).Value)
                                det.gprecio_unitario2 = Convert.ToDouble(DataGridView1.Rows(i1).Cells(6).Value / TextBox11.Text)
                                det.gventa_total2 = (DataGridView1.Rows(i1).Cells(9).Value / TextBox11.Text)
                                det.gvalor_venta2 = (DataGridView1.Rows(i1).Cells(7).Value / TextBox11.Text)
                                det.gigv2 = (DataGridView1.Rows(i1).Cells(8).Value / TextBox11.Text)
                                det.gtotal2 = (DataGridView1.Rows(i1).Cells(9).Value / TextBox11.Text)
                            Else
                                If RadioButton2.Checked = True Then
                                    det.gprecio_unitario = Convert.ToDouble(DataGridView1.Rows(i1).Cells(6).Value * TextBox11.Text)
                                    det.gventa_total = 0.00
                                    det.gvalor_venta = (DataGridView1.Rows(i1).Cells(7).Value * TextBox11.Text)
                                    det.gigv = (DataGridView1.Rows(i1).Cells(8).Value * TextBox11.Text)
                                    det.gtotal = (DataGridView1.Rows(i1).Cells(9).Value * TextBox11.Text)
                                    det.gprecio_unitario2 = Convert.ToDouble(DataGridView1.Rows(i1).Cells(6).Value)
                                    det.gventa_total2 = 0.00
                                    det.gvalor_venta2 = (DataGridView1.Rows(i1).Cells(7).Value)
                                    det.gigv2 = (DataGridView1.Rows(i1).Cells(8).Value)
                                    det.gtotal2 = (DataGridView1.Rows(i1).Cells(9).Value)
                                End If

                            End If

                            '------------------------
                            If Convert.ToString(DataGridView1.Rows(i1).Cells(10).Value) = "" Then
                                det.gop = ""
                            Else
                                det.gop = DataGridView1.Rows(i1).Cells(10).Value
                            End If

                            det.galmacen = TextBox10.Text
                            det.grubro = DataGridView1.Rows(i1).Cells(1).Value
                            det.gum = DataGridView1.Rows(i1).Cells(4).Value
                            det.ggrm = DataGridView1.Rows(i1).Cells(12).Value
                            det.gpartida = DataGridView1.Rows(i1).Cells(13).Value
                            det.gfatura = TextBox21.Text
                            det.gcorrelativo_fac = TextBox5.Text
                            det.gtido = TextBox31.Text
                            det.gano = Label27.Text
                            det.gccia = Label29.Text
                            det.grubrodetallado = Trim(DataGridView1.Rows(i1).Cells(14).Value)
                            cla.insertar_detalle_venta(det)

                            '----eliminar sinonimo
                            Dim hg As New vfacturasistema
                            hg.gcodigo_sin = Trim(DataGridView1.Rows(i1).Cells(2).Value)
                            hg.gccia = Label29.Text
                            cla.eliminar_sinonimo(hg)
                            ''----insertar sinonimo
                            Dim hj As New vfacturasistema
                            hj.gcodigo_sin = Trim(DataGridView1.Rows(i1).Cells(2).Value)
                            hj.gitem_sin = "001"
                            hj.gnomb_sin = DataGridView1.Rows(i1).Cells(3).Value
                            hj.gccia = Label29.Text
                            cla.insertar_sinonio(hj)

                        Next
                        ' insertra detalle contable
                        '12121101
                        Dim tt As New vdeta_contable
                        Dim ju56 As New vcontable
                        tt.gccia_3a = Label29.Text
                        tt.gcperiodo_3a = Label27.Text
                        ju56.gALMACEN = TextBox10.Text
                        ju56.gccia_3 = Label29.Text
                        Dim sumi2, sumi5 As Double
                        tt.gccom_3a = lp1.BUSCAR_SUBDIARIO(ju56)


                        'If Label28.Text = 1 Then
                        '    Dim func As New ffacturavianny
                        '    Dim dts As New vfacturavianny
                        '    Dim np As Int16
                        '    Select Case TextBox2.Text.Length

                        '        Case "1" : TextBox2.Text = "0" & "" & TextBox2.Text
                        '        Case "2" : TextBox2.Text = TextBox2.Text

                        '    End Select
                        '    dts.gmes = TextBox2.Text
                        '    dts.galmacen = Label5.Text
                        '    dts.gano = Label27.Text
                        '    dts.gccia = Label29.Text
                        '    np = func.buscar_correlativo_venta(dts) + 1
                        '    Select Case np.ToString.Length

                        '        Case "1" : tt.gncom_3a = TextBox2.Text & "" & "00000" & "" & np
                        '        Case "2" : tt.gncom_3a = TextBox2.Text & "" & "0000" & "" & np
                        '        Case "3" : tt.gncom_3a = TextBox2.Text & "" & "000" & "" & np
                        '        Case "4" : tt.gncom_3a = TextBox2.Text & "" & "00" & "" & np
                        '        Case "5" : tt.gncom_3a = TextBox2.Text & "" & "0" & "" & np
                        '        Case "6" : tt.gncom_3a = TextBox2.Text & "" & np
                        '    End Select

                        'Else
                        '    tt.gncom_3a = TextBox2.Text & "" & TextBox1.Text
                        'End If
                        tt.gncom_3a = cab.gncom
                        If RadioButton2.Checked = True Then
                            If TextBox20.Text = "69" Then
                                tt.gcuen_3a = "12121106 "
                            Else
                                tt.gcuen_3a = "12121102"
                            End If

                        Else
                            If TextBox20.Text = "69" Then
                                tt.gcuen_3a = "12121105"
                            Else
                                tt.gcuen_3a = "12121101"
                            End If

                        End If

                        tt.gcuend_3a = ""
                        tt.gcuenc_3a = ""
                        tt.gcodh_3a = "D"
                        tt.gfich_3a = TextBox6.Text
                        tt.gcdoc_3a = TextBox31.Text
                        tt.gndoc_3a = TextBox21.Text & "-" & TextBox5.Text
                        tt.gcref_3a = ""
                        tt.gnref_3a = ""
                        tt.glinea_3a = ""
                        tt.gpptof_3a = ""
                        tt.gnomb_3a = TextBox7.Text
                        tt.gglosa_3a = ""
                        If RadioButton1.Checked = True Then
                            For i = 0 To ad - 1
                                sumi2 = sumi2 + Val(DataGridView1.Rows(i).Cells(9).Value)
                            Next
                            sumi5 = sumi2 / tcambio
                            tt.gimp1_3a = sumi2
                            tt.gimp2_3a = sumi5
                        Else
                            If RadioButton2.Checked = True Then
                                For i = 0 To ad - 1
                                    sumi2 = sumi2 + Val(DataGridView1.Rows(i).Cells(9).Value)
                                Next
                                sumi5 = sumi2 * tcambio
                                tt.gimp1_3a = sumi5
                                tt.gimp2_3a = sumi2

                            End If
                        End If



                        tt.gfref_3a = DateTimePicker2.Value
                        tt.gflagd_3a = 0
                        tt.gproy_3a = ""
                        tt.gprinci_3a = ""
                        tt.gfuen_3a = ""
                        tt.gparti_3a = ""
                        tt.gtefe_3a = ""
                        tt.gsitlet_3a = ""
                        tt.gbanco_3a = ""
                        tt.gnrounico_3a = ""
                        tt.gendoso_3a = ""
                        tt.gaval_3a = ""
                        tt.gflag_3a = 1
                        tt.gocompra_3a = ""
                        tt.gordens_3a = ""
                        tt.gtcam_3a = TextBox11.Text
                        tt.gfemision_3a = DateTimePicker1.Value
                        lp1.insertar_detallecontable(tt)
                        '40111101     
                        Dim tt1 As New vdeta_contable
                        Dim ju561 As New vcontable
                        tt1.gccia_3a = Label29.Text
                        tt1.gcperiodo_3a = Label27.Text
                        ju561.gALMACEN = TextBox10.Text
                        ju561.gccia_3 = Label29.Text
                        Dim sume1, sume4 As Double
                        tt1.gccom_3a = lp1.BUSCAR_SUBDIARIO(ju561)


                        'If Label28.Text = 1 Then
                        '    Dim func As New ffacturavianny
                        '    Dim dts As New vfacturavianny
                        '    Dim np As Int16
                        '    Select Case TextBox2.Text.Length

                        '        Case "1" : TextBox2.Text = "0" & "" & TextBox2.Text
                        '        Case "2" : TextBox2.Text = TextBox2.Text

                        '    End Select
                        '    dts.gmes = TextBox2.Text
                        '    dts.galmacen = Label5.Text
                        '    dts.gano = Label27.Text
                        '    dts.gccia = Label29.Text
                        '    np = func.buscar_correlativo_venta(dts) + 1
                        '    Select Case np.ToString.Length

                        '        Case "1" : tt1.gncom_3a = TextBox2.Text & "" & "00000" & "" & np
                        '        Case "2" : tt1.gncom_3a = TextBox2.Text & "" & "0000" & "" & np
                        '        Case "3" : tt1.gncom_3a = TextBox2.Text & "" & "000" & "" & np
                        '        Case "4" : tt1.gncom_3a = TextBox2.Text & "" & "00" & "" & np
                        '        Case "5" : tt1.gncom_3a = TextBox2.Text & "" & "0" & "" & np
                        '        Case "6" : tt1.gncom_3a = TextBox2.Text & "" & np
                        '    End Select

                        'Else
                        '    tt1.gncom_3a = TextBox2.Text & "" & TextBox1.Text
                        'End If
                        tt1.gncom_3a = cab.gncom
                        tt1.gcuen_3a = "40111101"
                        tt1.gcuend_3a = ""
                        tt1.gcuenc_3a = ""
                        tt1.gcodh_3a = "H"
                        tt1.gfich_3a = TextBox6.Text
                        tt1.gcdoc_3a = TextBox31.Text
                        tt1.gndoc_3a = TextBox21.Text & "-" & TextBox5.Text
                        tt1.gcref_3a = ""
                        tt1.gnref_3a = ""
                        tt1.glinea_3a = ""
                        tt1.gpptof_3a = ""
                        tt1.gnomb_3a = TextBox7.Text
                        tt1.gglosa_3a = ""

                        If RadioButton1.Checked = True Then
                            For i = 0 To ad - 1


                                sume1 = sume1 + Val(DataGridView1.Rows(i).Cells(8).Value)

                            Next

                            sume4 = sume1 / tcambio

                            tt1.gimp1_3a = sume1
                            tt1.gimp2_3a = sume4

                        Else
                            If RadioButton2.Checked = True Then
                                For i = 0 To ad - 1


                                    sume1 = sume1 + Val(DataGridView1.Rows(i).Cells(8).Value)

                                Next

                                sume4 = sume1 * tcambio

                                tt1.gimp1_3a = sume4
                                tt1.gimp2_3a = sume1


                            End If
                        End If



                        tt1.gfref_3a = DateTimePicker2.Value
                        tt1.gflagd_3a = 0
                        tt1.gproy_3a = ""
                        tt1.gprinci_3a = ""
                        tt1.gfuen_3a = ""
                        tt1.gparti_3a = ""
                        tt1.gtefe_3a = ""
                        tt1.gsitlet_3a = ""
                        tt1.gbanco_3a = ""
                        tt1.gnrounico_3a = ""
                        tt1.gendoso_3a = ""
                        tt1.gaval_3a = ""
                        tt1.gflag_3a = 1
                        tt1.gocompra_3a = ""
                        tt1.gordens_3a = ""
                        tt1.gtcam_3a = TextBox11.Text
                        tt1.gfemision_3a = DateTimePicker1.Value
                        lp1.insertar_detallecontable(tt1)
                        'DE ACUERDO AL RUBRO
                        Dim tt2 As New vdeta_contable
                        Dim tt3 As New vcontable
                        Dim ju562 As New vcontable
                        tt2.gccia_3a = Label29.Text
                        tt2.gcperiodo_3a = Label27.Text
                        ju562.gALMACEN = TextBox10.Text
                        ju562.gccia_3 = Label29.Text
                        Dim sumee, sumee1, sumee2, sumee3, sumee4, sumee5 As Double
                        tt2.gccom_3a = lp1.BUSCAR_SUBDIARIO(ju562)

                        tt2.gncom_3a = cab.gncom
                        tt3.gRUBRO = DataGridView1.Rows(0).Cells(1).Value
                        tt3.gccia_3 = Label29.Text
                        tt2.gcuen_3a = lp1.BUSCAR_RUBRO(tt3)
                        tt2.gcuend_3a = ""
                        tt2.gcuenc_3a = ""
                        tt2.gcodh_3a = "H"
                        tt2.gfich_3a = TextBox6.Text
                        tt2.gcdoc_3a = TextBox31.Text
                        tt2.gndoc_3a = TextBox21.Text & "-" & TextBox5.Text
                        tt2.gcref_3a = ""
                        tt2.gnref_3a = ""
                        tt2.glinea_3a = ""
                        tt2.gpptof_3a = ""
                        tt2.gnomb_3a = TextBox7.Text
                        tt2.gglosa_3a = ""
                        If RadioButton1.Checked = True Then
                            For i = 0 To ad - 1

                                sumee = sumee + Val(DataGridView1.Rows(i).Cells(7).Value)
                                sumee1 = sumee1 + Val(DataGridView1.Rows(i).Cells(8).Value)
                                sumee2 = sumee2 + Val(DataGridView1.Rows(i).Cells(9).Value)
                            Next
                            sumee3 = sumee / tcambio
                            sumee4 = sumee1 / tcambio
                            sumee5 = sumee2 / tcambio


                            tt2.gimp1_3a = sumee


                            tt2.gimp2_3a = sumee3

                        Else
                            If RadioButton2.Checked = True Then
                                For i = 0 To ad - 1

                                    sumee = sumee + Val(DataGridView1.Rows(i).Cells(7).Value)
                                    sumee1 = sumee1 + Val(DataGridView1.Rows(i).Cells(8).Value)
                                    sumee2 = sumee2 + Val(DataGridView1.Rows(i).Cells(9).Value)
                                Next
                                sumee3 = sumee * tcambio
                                sumee4 = sumee1 * tcambio
                                sumee5 = sumee2 * tcambio

                                tt2.gimp1_3a = sumee3


                                tt2.gimp2_3a = sumee

                            End If
                        End If




                        tt2.gfref_3a = DateTimePicker2.Value
                        tt2.gflagd_3a = 0
                        tt2.gproy_3a = ""
                        tt2.gprinci_3a = ""
                        tt2.gfuen_3a = ""
                        tt2.gparti_3a = ""
                        tt2.gtefe_3a = ""
                        tt2.gsitlet_3a = ""
                        tt2.gbanco_3a = ""
                        tt2.gnrounico_3a = ""
                        tt2.gendoso_3a = ""
                        tt2.gaval_3a = ""
                        tt2.gflag_3a = 1
                        tt2.gocompra_3a = ""
                        tt2.gordens_3a = ""
                        tt2.gtcam_3a = TextBox11.Text
                        tt2.gfemision_3a = DateTimePicker1.Value
                        lp1.insertar_detallecontable(tt2)
                        'ENVIAR_CORREO()
                        MsgBox("LA VENTA SE REGISTRO CON EXITO")
                        Dim hj2 As New flog
                        Dim hj1 As New vlog
                        hj1.gmodulo = "VENTAS_FACTURA"
                        hj1.gcuser = MDIParent1.Label3.Text
                        Select Case Label28.Text
                            Case "1" : hj1.gaccion = 1
                            Case "2" : hj1.gaccion = 2
                            Case "3" : hj1.gaccion = 3

                        End Select

                        hj1.gpc = My.Computer.Name
                        hj1.gfecha = String.Format("{0:G}", DateTime.Now)
                        hj1.gdetalle = TextBox21.Text & TextBox5.Text
                        hj1.gccia = Label29.Text
                        hj2.insertar_log(hj1)
                        mostrar_corrlativo()
                        TextBox21.Text = ""
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox7.Text = ""
                        TextBox23.Text = ""
                        TextBox13.Text = ""
                        TextBox15.Text = ""
                        TextBox22.Text = ""
                        TextBox25.Text = ""
                        TextBox26.Text = ""
                        TextBox28.Text = ""
                        TextBox29.Text = ""
                        TextBox14.Text = ""
                        TextBox27.Text = ""
                        TextBox30.Text = ""
                        TextBox33.Text = ""
                        TextBox34.Text = ""
                        ComboBox2.Text = ""
                        DataGridView1.Rows.Clear()
                        TextBox21.Enabled = False
                        TextBox5.Enabled = False
                        TextBox6.Enabled = False
                        TextBox7.Enabled = False
                        TextBox23.Enabled = False
                        TextBox13.Enabled = False
                        TextBox15.Enabled = False
                        TextBox22.Enabled = False
                        TextBox25.Enabled = False
                        TextBox26.Enabled = False
                        TextBox28.Enabled = False
                        TextBox29.Enabled = False
                        TextBox14.Enabled = False
                        ComboBox1.SelectedIndex = 0
                        DataGridView1.Enabled = False
                    End If
                End If
            End If
        End If
    End Sub




    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        TextBox16.Text = "04"
        TextBox17.Text = "VENTAS LOCALES - SOLES"
        RadioButton4.Checked = False
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        TextBox16.Text = "32"
        TextBox17.Text = "FACTURAS/BOLETAS SIN VALOR COMERCIAL"
        RadioButton3.Checked = False
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Pedidos.Label3.Text = Label27.Text
        Pedidos.Label4.Text = 1
        Pedidos.Label5.Text = Label29.Text
        Pedidos.Label6.Text = TextBox14.Text
        Pedidos.Label7.Text = TextBox6.Text
        Pedidos.Show()
    End Sub

    Dim dt6 As New DataTable

    Private Sub PictureBox2_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub GroupBox12_Enter(sender As Object, e As EventArgs) Handles GroupBox12.Enter

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        MsgBox("LA OPCION ESTA DESHABILITADA, SOLICITAR AL AREA CONTABLE LA ANULACION")

        'Dim va As String
        'If Label28.Text = 1 Then
        '    Dim func As New ffacturavianny
        '    Dim dts As New vfacturavianny
        '    Dim np As Integer

        '    Select Case TextBox2.Text.Length

        '        Case "1" : TextBox2.Text = "0" & "" & TextBox2.Text
        '        Case "2" : TextBox2.Text = TextBox2.Text

        '    End Select
        '    dts.gmes = TextBox2.Text
        '    dts.galmacen = Label5.Text
        '    dts.gano = Label27.Text
        '    dts.gccia = Label29.Text
        '    np = func.buscar_correlativo_venta(dts) + 1
        '    Select Case np.ToString.Length

        '        Case "1" : va = "00000" & "" & np
        '        Case "2" : va = "0000" & "" & np
        '        Case "3" : va = "000" & "" & np
        '        Case "4" : va = "00" & "" & np
        '        Case "5" : va = "0" & "" & np
        '        Case "6" : va = np
        '    End Select

        'Else
        '    va = TextBox2.Text & "" & TextBox1.Text
        'End If
        'MsgBox(va)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        DataGridView1.Rows.Clear()
        TextBox25.Text = ""
        TextBox26.Text = ""
        TextBox28.Text = ""
        TextBox29.Text = ""
        TextBox34.Text = ""
        TextBox13.Text = ""
        TextBox23.Text = ""
        TextBox14.Text = ""
        TextBox22.Text = ""
        TextBox7.Text = ""
        TextBox6.Text = ""
        TextBox5.Text = ""
        TextBox21.Text = ""
        TextBox33.Text = ""
        TextBox30.Text = ""
        TextBox27.Text = ""


        TextBox21.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox23.Enabled = False
        TextBox13.Enabled = False
        TextBox15.Enabled = False
        TextBox22.Enabled = False
        TextBox25.Enabled = False
        TextBox26.Enabled = False
        TextBox28.Enabled = False
        TextBox29.Enabled = False
        TextBox14.Enabled = False
        ComboBox1.SelectedIndex = 0
        DataGridView1.Enabled = False
        TextBox1.Text = ""
        TextBox2.Select()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Button4.Enabled = True
        TextBox13.Enabled = True
        TextBox22.Enabled = True
        TextBox22.ReadOnly = False
        DateTimePicker1.Enabled = True
        DateTimePicker2.Enabled = True
        TextBox6.Enabled = True
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        DataGridView1.Enabled = True
        DataGridView1.ReadOnly = False
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        RadioButton4.Enabled = True
        Label28.Text = 2
        TextBox18.Enabled = True
        TextBox19.Enabled = True
        TextBox18.ReadOnly = False
        TextBox19.ReadOnly = False
        Button6.Enabled = True
        Button3.Enabled = True
        Button5.Enabled = True
        Button8.Enabled = True

    End Sub

    Private Sub TextBox33_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox33.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim ab As New vfacturasistema
            Dim ab1 As New ffacturavianny
            Select Case TextBox33.Text.Length

                Case "1" : TextBox33.Text = "0000000" & "" & TextBox33.Text
                Case "2" : TextBox33.Text = "000000" & "" & TextBox33.Text
                Case "3" : TextBox33.Text = "00000" & "" & TextBox33.Text
                Case "4" : TextBox33.Text = "0000" & "" & TextBox33.Text
                Case "5" : TextBox33.Text = "000" & "" & TextBox33.Text
                Case "6" : TextBox33.Text = "00" & "" & TextBox33.Text
                Case "7" : TextBox33.Text = "0" & "" & TextBox33.Text
                Case "8" : TextBox33.Text = TextBox33.Text
            End Select
            ab.gtdoc = TextBox27.Text
            ab.gfatura = TextBox30.Text
            ab.gcorrelativo_fac = TextBox33.Text
            ab.gccia = Label29.Text
            dt6 = ab1.mostrar_factura_notacredito(ab)
            If dt6.Rows.Count > 0 Then
                Dim JI As Integer
                DataGridView6.DataSource = dt6
                TextBox6.Text = DataGridView6.Rows(0).Cells(0).Value
                TextBox7.Text = DataGridView6.Rows(0).Cells(1).Value
                TextBox23.Text = DataGridView6.Rows(0).Cells(3).Value
                TextBox13.Text = DataGridView6.Rows(0).Cells(4).Value
                TextBox14.Text = Trim(DataGridView6.Rows(0).Cells(5).Value)
                TextBox20.Text = DataGridView6.Rows(0).Cells(2).Value
                TextBox24.Text = DataGridView6.Rows(0).Cells(6).Value
                abrir()
                enunciado3 = New SqlCommand("select alias_ven from custom_vianny.dbo.Vendedores where codigo_ven = " + DataGridView6.Rows(0).Cells(5).Value, conx)

                respuesta3 = enunciado3.ExecuteReader

                While respuesta3.Read

                    ComboBox6.Items.Clear()
                    ComboBox6.Items.Add(respuesta3.Item("alias_ven"))
                    ComboBox6.SelectedIndex = 0
                End While
                respuesta3.Close()
                If DataGridView6.Rows(0).Cells(19).Value = 1 Then
                    RadioButton1.Checked = True
                Else
                    RadioButton2.Checked = True
                End If
                CheckBox1.Checked = True
                CheckBox2.Checked = True
                Button7.Visible = True
                JI = DataGridView6.Rows.Count
                DataGridView1.Rows.Add(JI)
                For I = 0 To JI - 1
                    DataGridView1.Rows(I).Cells(0).Value = DataGridView6.Rows(I).Cells(7).Value
                    DataGridView1.Rows(I).Cells(1).Value = "0038"
                    DataGridView1.Rows(I).Cells(2).Value = DataGridView6.Rows(I).Cells(8).Value
                    DataGridView1.Rows(I).Cells(3).Value = DataGridView6.Rows(I).Cells(9).Value
                    DataGridView1.Rows(I).Cells(4).Value = DataGridView6.Rows(I).Cells(10).Value
                    DataGridView1.Rows(I).Cells(5).Value = DataGridView6.Rows(I).Cells(11).Value
                    DataGridView1.Rows(I).Cells(6).Value = DataGridView6.Rows(I).Cells(12).Value
                    DataGridView1.Rows(I).Cells(7).Value = DataGridView6.Rows(I).Cells(13).Value
                    DataGridView1.Rows(I).Cells(8).Value = DataGridView6.Rows(I).Cells(14).Value
                    DataGridView1.Rows(I).Cells(9).Value = DataGridView6.Rows(I).Cells(15).Value
                    DataGridView1.Rows(I).Cells(11).Value = DataGridView6.Rows(I).Cells(16).Value
                    DataGridView1.Rows(I).Cells(12).Value = DataGridView6.Rows(I).Cells(17).Value
                    DataGridView1.Rows(I).Cells(13).Value = DataGridView6.Rows(I).Cells(18).Value
                Next
            Else
                MsgBox("LA FACTURA INGRESADA NO EXISTE")
            End If
            TextBox18.ReadOnly = True
            TextBox19.ReadOnly = True
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged

    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox5.Text.Length

                Case "1" : TextBox5.Text = "0000000" & "" & TextBox5.Text
                Case "2" : TextBox5.Text = "000000" & "" & TextBox5.Text
                Case "3" : TextBox5.Text = "00000" & "" & TextBox5.Text
                Case "4" : TextBox5.Text = "0000" & "" & TextBox5.Text
                Case "5" : TextBox5.Text = "000" & "" & TextBox5.Text
                Case "6" : TextBox5.Text = "00" & "" & TextBox5.Text
                Case "7" : TextBox5.Text = "0" & "" & TextBox5.Text
                Case "8" : TextBox5.Text = TextBox5.Text
            End Select
            Dim sql1021011 As String = "select * from custom_vianny.dbo.fag0300 where ccia_3='" + Label29.Text + "'  and sfactu_3 ='" + TextBox21.Text + "' and nfactu_3 ='" + TextBox5.Text + "' and tidoc_3 ='001'"
            Dim cmd1021011 As New SqlCommand(sql1021011, conx)
            Rsr21 = cmd1021011.ExecuteReader()
            If Rsr21.Read() Then
                If Rsr21(0) > 0 Then
                    TextBox18.Enabled = False
                    TextBox19.Enabled = False
                    MsgBox("EL CORRELATIVO DE LA FACTURA YA EXISTE INGRESE UN CORRELATIVO VACIO")

                End If
            Else

                TextBox18.Enabled = True
                TextBox19.Enabled = True
                TextBox18.Select()
            End If
            Rsr21.Close()

        End If
    End Sub

    Private Sub DateTimePicker1_Click(sender As Object, e As EventArgs) Handles DateTimePicker1.Click

    End Sub

    Private Sub TextBox1_DoubleClick(sender As Object, e As EventArgs) Handles TextBox1.DoubleClick
        Try
            Dim func As New ffacturavianny
            Dim dts As New vfacturavianny

            Select Case TextBox2.Text.Length

                Case "1" : TextBox2.Text = "0" & "" & TextBox2.Text
                Case "2" : TextBox2.Text = TextBox2.Text

            End Select
            dts.gmes = TextBox2.Text
            dts.galmacen = Label5.Text
            dts.gano = Label27.Text
            dts.gccia = Label29.Text
            Me.TextBox1.Text = func.buscar_correlativo_venta(dts) + 1
            Select Case TextBox1.Text.Length

                Case "1" : TextBox1.Text = "00000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "0000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "6" : TextBox1.Text = TextBox1.Text
            End Select
            ComboBox1.Enabled = True
            TextBox21.ReadOnly = False
            ComboBox1.Select()

            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox23.Text = ""
            TextBox13.Text = ""
            TextBox15.Text = ""
            TextBox22.Text = ""
            TextBox25.Text = ""
            TextBox26.Text = ""
            TextBox28.Text = ""
            TextBox29.Text = ""
            TextBox14.Text = ""
            TextBox27.Text = ""
            TextBox30.Text = ""
            TextBox33.Text = ""
            TextBox34.Text = ""
            ComboBox2.Text = ""
            DataGridView1.Rows.Clear()


            TextBox6.Enabled = False
            TextBox7.Enabled = False
            TextBox23.Enabled = False
            TextBox13.Enabled = False
            TextBox15.Enabled = False
            TextBox22.Enabled = False
            TextBox25.Enabled = False
            TextBox26.Enabled = False
            TextBox28.Enabled = False
            TextBox29.Enabled = False
            TextBox14.Enabled = False
            ComboBox1.SelectedIndex = 0
            DataGridView1.Enabled = False
        Catch ex As Exception



        End Try
    End Sub

    Private Sub TextBox1_Invalidated(sender As Object, e As InvalidateEventArgs) Handles TextBox1.Invalidated

    End Sub

    Private Sub TextBox19_MouseWheel(sender As Object, e As MouseEventArgs) Handles TextBox19.MouseWheel

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress

    End Sub
End Class