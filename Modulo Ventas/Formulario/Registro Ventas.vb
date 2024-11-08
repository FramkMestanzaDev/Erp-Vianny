Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Imports System.Data.SqlClient

Public Class Registro_Ventas
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE rpm_ven ='2' AND ccia_ven ='" + Label26.Text + "'", conx)
            conn.Fill(ty2)
            ComboBox6.DataSource = ty2
            ComboBox6.DisplayMember = "alias_ven"
            ComboBox6.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Registro_Venta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ty2.Clear()
        ComboBox1.SelectedIndex = 0
        Dim func As New ftcambio
        Dim dts As New vtcambio
        dts.gfecha = DateTimePicker1.Text
        TextBox11.Text = func.mostrar_tipo_cambio(dts)
        RadioButton1.Checked = True
        RadioButton3.Checked = True
        mostrar_correlativo_venta()
        TextBox3.Select()
        Label22.Text = 1
        abrir()
        llenar_combo_box2()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Label26.Text = "01" Then
            TextBox31.Text = "001"
            TextBox21.Text = "N015"
        Else
            TextBox31.Text = "001"
            TextBox21.Text = "N016"
        End If


        If TextBox21.Text = "" Then

            MsgBox("FALTA INGRESAR LA SERIE DE LA FACTURA")
        Else
            Dim func As New vdocfac
            Dim func1 As New fbolfac
            func.gdoc = TextBox31.Text
            func.gccia = Label26.Text

            TextBox5.Text = func1.mostrar_correlativo(func) + 1

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
            TextBox5.Select()
        End If
    End Sub
    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim bo As New fregistro_venta
            Dim bo1 As New vregistroventa
            Select Case TextBox6.Text.Length
                Case "1" : TextBox6.Text = TextBox6.Text & "" & "0000000000"
                Case "2" : TextBox6.Text = TextBox6.Text & "" & "000000000"
                Case "3" : TextBox6.Text = TextBox6.Text & "" & "00000000"
                Case "4" : TextBox6.Text = TextBox6.Text & "" & "0000000"
                Case "5" : TextBox6.Text = TextBox6.Text & "" & "000000"
                Case "6" : TextBox6.Text = TextBox6.Text & "" & "00000"
                Case "7" : TextBox6.Text = TextBox6.Text & "" & "0000"
                Case "8" : TextBox6.Text = TextBox6.Text & "" & "000"
                Case "9" : TextBox6.Text = TextBox6.Text & "" & "000"
                Case "10" : TextBox6.Text = TextBox6.Text & "" & "00"
                Case "11" : TextBox6.Text = TextBox6.Text & "" & "0"
            End Select
            bo1.gruc = TextBox6.Text
            TextBox7.Text = bo.motrar_rsocial(bo1)
        End If
    End Sub
    Private Sub mostrar_correlativo_venta()
        Try
            Dim func As New fgu_comprobante
            Dim dts As New vnia


            dts.gmes = DateTime.Now.ToString("MM")
            dts.gano = MDIParent1.Label5.Text
            TextBox4.Text = DateTime.Now.ToString("MM")
            dts.gccia = Label26.Text
            Me.TextBox3.Text = func.buscar_coprobante(dts) + 1

            Select Case TextBox3.Text.Length

                Case "1" : TextBox3.Text = "00000" & "" & TextBox3.Text
                Case "2" : TextBox3.Text = "0000" & "" & TextBox3.Text
                Case "3" : TextBox3.Text = "000" & "" & TextBox3.Text
                Case "4" : TextBox3.Text = "00" & "" & TextBox3.Text
                Case "5" : TextBox3.Text = "0" & "" & TextBox3.Text
                Case "6" : TextBox3.Text = TextBox3.Text
            End Select

        Catch ex As Exception
            MsgBox(ex.Message)


        End Try
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim func As New ftcambio
        Dim dts As New vtcambio
        dts.gfecha = DateTimePicker1.Text
        TextBox11.Text = func.mostrar_tipo_cambio(dts)
    End Sub

    Dim RT, RT2 As New DataTable
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        Dim valo As String
        Dim i As Integer
        Dim FG As New fregistro_venta
        Dim JK As New vdetalleregistro
        Dim HJ As Integer
        valo = TextBox3.Text
        DataGridView1.Rows.Clear()

        If e.KeyCode = Keys.Enter Then


            i = Len(valo)
            Select Case i
                Case "1" : valo = "00000" & "" & valo
                Case "2" : valo = "0000" & "" & valo
                Case "3" : valo = "000" & "" & valo
                Case "4" : valo = "00" & "" & valo
                Case "5" : valo = "0" & "" & valo
                Case "6" : valo = valo

            End Select
            TextBox3.Text = valo
            JK.gncom_3a = TextBox4.Text & "" & TextBox3.Text
            JK.gcperiodo_3a = Label17.Text
            JK.gempresa = Label26.Text
            HJ = FG.existeregistrovn(JK)
            If HJ > 0 Then
                'DataGridView1.Rows.Clear()

                RT = FG.buscar_ventas_negra_DETALLE(JK)
                DataGridView9.DataSource = RT
                Dim UI As Integer
                UI = DataGridView9.Rows.Count
                DataGridView1.Rows.Add(UI)
                Dim SUMA, SUMA1, SUMA2, SUMA3 As Double
                For i = 0 To UI - 1
                    DataGridView1.Rows(i).Cells(0).Value = DataGridView9.Rows(i).Cells(0).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView9.Rows(i).Cells(1).Value
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView9.Rows(i).Cells(2).Value
                    DataGridView1.Rows(i).Cells(3).Value = DataGridView9.Rows(i).Cells(3).Value
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView9.Rows(i).Cells(4).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView9.Rows(i).Cells(5).Value
                    DataGridView1.Rows(i).Cells(6).Value = DataGridView9.Rows(i).Cells(6).Value
                    DataGridView1.Rows(i).Cells(7).Value = DataGridView9.Rows(i).Cells(7).Value
                    DataGridView1.Rows(i).Cells(8).Value = DataGridView9.Rows(i).Cells(8).Value
                    DataGridView1.Rows(i).Cells(9).Value = DataGridView9.Rows(i).Cells(9).Value
                    DataGridView1.Rows(i).Cells(10).Value = DataGridView9.Rows(i).Cells(10).Value
                    DataGridView1.Rows(i).Cells(11).Value = DataGridView9.Rows(i).Cells(11).Value
                    DataGridView1.Rows(i).Cells(12).Value = DataGridView9.Rows(i).Cells(12).Value
                    SUMA = SUMA + DataGridView9.Rows(i).Cells(4).Value
                    SUMA1 = SUMA1 + DataGridView9.Rows(i).Cells(6).Value
                    SUMA2 = SUMA2 + DataGridView9.Rows(i).Cells(7).Value
                    SUMA3 = SUMA3 + DataGridView9.Rows(i).Cells(8).Value
                Next
                TextBox25.Text = SUMA
                TextBox26.Text = SUMA1
                TextBox28.Text = SUMA2
                TextBox29.Text = SUMA3
                RT2 = FG.buscar_ventas_negra_cabecera(JK)
                DataGridView10.DataSource = RT2
                TextBox31.Text = DataGridView10.Rows(0).Cells(0).Value
                TextBox21.Text = DataGridView10.Rows(0).Cells(1).Value
                TextBox5.Text = DataGridView10.Rows(0).Cells(2).Value
                TextBox34.Text = DataGridView10.Rows(0).Cells(16).Value
                TextBox6.Text = DataGridView10.Rows(0).Cells(3).Value
                TextBox7.Text = DataGridView10.Rows(0).Cells(4).Value
                DateTimePicker1.Value = DataGridView10.Rows(0).Cells(5).Value
                TextBox11.Text = DataGridView10.Rows(0).Cells(6).Value

                Select Case DataGridView10.Rows(0).Cells(7).Value
                    Case "01" : ComboBox5.SelectedIndex = 0
                    Case "02" : ComboBox5.SelectedIndex = 1
                    Case "03" : ComboBox5.SelectedIndex = 2
                    Case "10" : ComboBox5.SelectedIndex = 3
                    Case "04" : ComboBox5.SelectedIndex = 4
                    Case "05" : ComboBox5.SelectedIndex = 5
                    Case "06" : ComboBox5.SelectedIndex = 6
                    Case "07" : ComboBox5.SelectedIndex = 7
                    Case "08" : ComboBox5.SelectedIndex = 8
                    Case "09" : ComboBox5.SelectedIndex = 9

                End Select
                TextBox12.Text = DataGridView10.Rows(0).Cells(8).Value
                TextBox13.Text = DataGridView10.Rows(0).Cells(9).Value
                'TextBox17.Text = DataGridView10.Rows(0).Cells(10).Value
                If DataGridView10.Rows(0).Cells(11).Value = 1 Then
                    RadioButton1.Checked = True
                Else
                    RadioButton2.Checked = True
                End If
                If DataGridView10.Rows(0).Cells(12).Value = 1 Then
                    RadioButton3.Checked = True
                Else
                    RadioButton4.Checked = True
                End If
                If DataGridView10.Rows(0).Cells(13).Value = 1 Then
                    CheckBox1.Checked = True
                End If
                If DataGridView10.Rows(0).Cells(14).Value = 1 Then
                    CheckBox2.Checked = True
                End If

                abrir()
                Dim sql102 As String = "SELECT  alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  codigo_ven ='" + DataGridView10.Rows(0).Cells(15).Value + "'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr3 = cmd102.ExecuteReader()
                If Rsr3.Read() = True Then
                    ComboBox6.Text = Rsr3(0)

                End If

                Rsr3.Close()
                'Select Case DataGridView10.Rows(0).Cells(15).Value
                '    Case "0010" : ComboBox6.Text = "GBEDON"
                '    Case "0022" : ComboBox6.Text = "VINCIO"
                '    Case "0023" : ComboBox6.Text = "DBRAVO"
                '    Case "0025" : ComboBox6.Text = "JSALINAS"
                '    Case "0029" : ComboBox6.Text = "GCUEVA"
                '    Case "0026" : ComboBox6.Text = "AMENDO"
                '    Case "0007" : ComboBox6.Text = "VGRAUS"
                '    Case "0027" : ComboBox6.Text = "VPIZARRO"
                '    Case "0028" : ComboBox6.Text = "JBALVIN"
                '    Case "0005" : ComboBox6.Text = "VSILVERIO"
                '    Case "0034" : ComboBox6.Text = "WSALINAS"
                '    Case "0037" : ComboBox6.Text = "MGRAUS"
                'End Select
                DataGridView9.DataSource = ""
                DataGridView10.DataSource = ""
                PictureBox1.Visible = True
                PictureBox4.Visible = True
                Button5.Enabled = False
            Else
                TextBox5.Select()
                'mostrar_correlativo_venta()

                Dim func As New ftcambio
                Dim dts As New vtcambio
                dts.gfecha = DateTimePicker1.Text
                TextBox11.Text = func.mostrar_tipo_cambio(dts)
                RadioButton1.Checked = True
                RadioButton3.Checked = True
                'mostrar_correlativo_venta()
                TextBox5.Select()
                'DataGridView1.Rows.Clear()
                DataGridView2.DataSource = ""
                DataGridView3.DataSource = ""
                DataGridView4.DataSource = ""
                DataGridView5.DataSource = ""
                DataGridView6.DataSource = ""
                DataGridView7.DataSource = ""
                DataGridView8.DataSource = ""
                TextBox6.Text = ""
                TextBox7.Text = ""

                TextBox15.Text = ""
                TextBox18.Text = ""
                TextBox19.Text = ""
                TextBox10.Text = ""
                TextBox9.Text = ""
                TextBox8.Text = ""
                TextBox25.Text = ""
                TextBox26.Text = ""
                TextBox28.Text = ""
                TextBox29.Text = ""
                TextBox22.Text = ""
                Dim func32 As New vdocfac
                Dim func132 As New fbolfac
                func32.gdoc = TextBox31.Text
                func32.gccia = Label26.Text

                TextBox5.Text = func132.mostrar_correlativo(func32) + 1

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
            End If
        Else
            If e.KeyCode = Keys.F1 Then
                BUSCAR_VENTA.TextBox2.Text = Label17.Text
                BUSCAR_VENTA.TextBox3.Text = TextBox16.Text
                BUSCAR_VENTA.Label1.Text = 1
                BUSCAR_VENTA.Show()
            End If
        End If

    End Sub

    'Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
    '    If e.KeyChar = Chr(Keys.Tab) Then

    '        MsgBox("tab")

    '    End If
    'End Sub
    Private dt As New DataTable
    Private dt1 As New DataTable
    Private Sub cabecera_guia()
        Try
            Dim func As New fguia2
            Dim dts As New vguia2


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
            func.gccia = Label26.Text

            dt = dts.consultar_cabecera_guia(func)
            If dt.Rows.Count <> 0 Then
                DataGridView2.DataSource = dt
                If TextBox6.Text = "" Then


                    TextBox6.Text = DataGridView2.Rows(0).Cells(4).Value
                    TextBox7.Text = DataGridView2.Rows(0).Cells(5).Value
                    DateTimePicker1.Text = DataGridView2.Rows(0).Cells(0).Value
                    Dim HG As New fcliente
                    Dim HG1 As New vcliente
                    Dim JH As String
                    HG1.gruc = Trim(DataGridView2.Rows(0).Cells(4).Value)
                    JH = HG.BUSCAR_VENDEDOR_CLIENTE(HG1)
                    TextBox27.Text = JH
                    abrir()
                    Dim sql102 As String = "SELECT  alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  codigo_ven ='" + JH + "' AND ccia_ven ='" + Label26.Text + "'"
                    Dim cmd102 As New SqlCommand(sql102, conx)
                    Rsr4 = cmd102.ExecuteReader()
                    If Rsr4.Read() = True Then
                        ComboBox6.Text = Rsr4(0)

                    End If

                    Rsr4.Close()
                    'Select Case JH
                    '    Case "0010" : ComboBox6.Text = "GBEDON"
                    '    Case "0022" : ComboBox6.Text = "VINCIO"
                    '    Case "0023" : ComboBox6.Text = "DBRAVO"
                    '    Case "0025" : ComboBox6.Text = "JSALINAS"
                    '    Case "0029" : ComboBox6.Text = "GCUEVA"
                    '    Case "0026" : ComboBox6.Text = "AMENDO"
                    '    Case "0007" : ComboBox6.Text = "VGRAUS"
                    '    Case "0027" : ComboBox6.Text = "VPIZARRO"
                    '    Case "0028" : ComboBox6.Text = "JBALVIN"
                    '    Case "0005" : ComboBox6.Text = "VSILVERIO"
                    '    Case "0030" : ComboBox6.Text = "RMEDINA"
                    '    Case "0034" : ComboBox6.Text = "WSALINAS"
                    '    Case "0037" : ComboBox6.Text = "MGRAUS"
                    'End Select
                    'ComboBox6.Enabled = False
                    If JH = "False" Then
                        ComboBox6.Enabled = True
                    End If
                    TextBox14.Text = "1"
                Else
                    If TextBox6.Text = DataGridView2.Rows(0).Cells(4).Value Then
                        TextBox6.Text = DataGridView2.Rows(0).Cells(4).Value
                        TextBox7.Text = DataGridView2.Rows(0).Cells(5).Value
                        DateTimePicker1.Text = DataGridView2.Rows(0).Cells(0).Value
                        TextBox14.Text = "1"
                    Else
                        TextBox14.Text = "2"
                        MsgBox("NO SE PUEDE AGREGRA GUIAS DE DIFERENTE CLIENTES")
                    End If
                End If
            End If
        Catch ex As Exception



        End Try
    End Sub
    Dim dt7, dt8 As DataTable
    Private Sub cabecera_nsa()
        Try
            Dim func As New vnia
            Dim dts As New fnsa


            func.galmacen = TextBox10.Text
            Select Case TextBox8.Text.Length

                Case "1" : TextBox8.Text = "00000" & "" & TextBox8.Text
                Case "2" : TextBox8.Text = "0000" & "" & TextBox8.Text
                Case "3" : TextBox8.Text = "000" & "" & TextBox8.Text
                Case "4" : TextBox8.Text = "00" & "" & TextBox8.Text
                Case "5" : TextBox8.Text = "0" & "" & TextBox8.Text
                Case "6" : TextBox8.Text = TextBox8.Text
            End Select
            Select Case TextBox9.Text.Length

                Case "1" : TextBox9.Text = "0" & "" & TextBox9.Text
                Case "2" : TextBox9.Text = TextBox9.Text

            End Select
            func.gncom = TextBox9.Text & "" & TextBox8.Text
            func.gano = Label17.Text
            func.gccia = Label26.Text
            dt7 = dts.mostrar_cabecera(func)
            If dt7.Rows.Count <> 0 Then

                DataGridView7.DataSource = dt7
                If TextBox6.Text = "" Then


                    TextBox6.Text = DataGridView7.Rows(0).Cells(4).Value
                    TextBox7.Text = DataGridView7.Rows(0).Cells(5).Value
                    DateTimePicker1.Text = DataGridView7.Rows(0).Cells(0).Value
                    Dim HG As New fcliente
                    Dim HG1 As New vcliente
                    Dim JH As String
                    HG1.gruc = Trim(DataGridView7.Rows(0).Cells(4).Value)

                    JH = HG.BUSCAR_VENDEDOR_CLIENTE(HG1)
                    TextBox27.Text = JH
                    abrir()
                    Dim sql102 As String = "SELECT  alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  codigo_ven ='" + JH + "' AND ccia_ven ='" + Label26.Text + "'"
                    Dim cmd102 As New SqlCommand(sql102, conx)
                    Rsr5 = cmd102.ExecuteReader()
                    If Rsr5.Read() = True Then
                        ComboBox6.Text = Rsr5(0)

                    End If

                    Rsr5.Close()
                    'Select Case JH
                    '    Case "0010" : ComboBox6.Text = "GBEDON"
                    '    Case "0022" : ComboBox6.Text = "VINCIO"
                    '    Case "0023" : ComboBox6.Text = "DBRAVO"
                    '    Case "0025" : ComboBox6.Text = "JSALINAS"
                    '    Case "0029" : ComboBox6.Text = "GCUEVA"
                    '    Case "0026" : ComboBox6.Text = "AMENDO"
                    '    Case "0007" : ComboBox6.Text = "VGRAUS"
                    '    Case "0027" : ComboBox6.Text = "VPIZARRO"
                    '    Case "0028" : ComboBox6.Text = "JBALVIN"
                    '    Case "0005" : ComboBox6.Text = "VSILVERIO"
                    '    Case "0030" : ComboBox6.Text = "RMEDINA"
                    '    Case "0034" : ComboBox6.Text = "WSALINAS"
                    '    Case "0037" : ComboBox6.Text = "MGRAUS"
                    'End Select
                    If JH = "False" Then
                        ComboBox6.Enabled = True
                    End If
                    TextBox14.Text = "1"
                Else
                    If TextBox6.Text = DataGridView7.Rows(0).Cells(4).Value Then
                        TextBox6.Text = DataGridView7.Rows(0).Cells(4).Value
                        TextBox7.Text = DataGridView7.Rows(0).Cells(5).Value
                        DateTimePicker1.Text = DataGridView7.Rows(0).Cells(0).Value
                        TextBox14.Text = "1"
                    Else
                        TextBox14.Text = "2"
                        MsgBox("NO SE PUEDE AGREGRA GUIAS DE DIFERENTE CLIENTES")
                    End If

                End If
            End If
        Catch ex As Exception



        End Try
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

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub TextBox19_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox19.KeyDown
        If e.KeyCode = Keys.Enter Then
            'DataGridView1.Rows.Clear()
            Dim NU As Integer
            Dim hj As New fventasn
            Dim hj1 As New vventas_n
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
            hj1.gccia = Label26.Text
            NU = hj.verificar_guia(hj1)

            If NU > 0 Then
                Dim respuesta As Integer

                respuesta = MsgBox("LA GUIA YA ESTA REGISTRADA, DESEA AGREGARLA?", vbQuestion + vbOKCancel)
                If respuesta = 1 Then '-- 2=Cancelar;1 = Aceptar
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
                    ComboBox5.Enabled = True
                    TextBox18.Text = ""
                    TextBox19.Text = ""
                    DataGridView1.Enabled = True
                    DataGridView1.ReadOnly = False
                    DataGridView1.Columns(0).ReadOnly = True
                    DataGridView1.Columns(4).ReadOnly = False
                    DataGridView1.Columns(5).ReadOnly = False
                    RadioButton1.Checked = True
                    CheckBox1.Checked = True
                    CheckBox2.Checked = True
                    TextBox22.ReadOnly = False
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
                ComboBox5.Enabled = True
                TextBox18.Text = ""
                TextBox19.Text = ""
                DataGridView1.Enabled = True
                DataGridView1.ReadOnly = False
                DataGridView1.Columns(0).ReadOnly = True
                DataGridView1.Columns(4).ReadOnly = False
                DataGridView1.Columns(5).ReadOnly = False
                RadioButton1.Checked = True
                CheckBox1.Checked = True
                CheckBox2.Checked = True
                TextBox22.ReadOnly = False
            End If


        End If

    End Sub


    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("Anexar Guia ? Si la respuesta es NO Anexara a una Nota de Salida", "COPIAR INFORMACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                TextBox18.Select()
                TextBox18.ReadOnly = False
                TextBox19.ReadOnly = False

            Else
                TextBox10.Select()
                TextBox8.ReadOnly = False
                TextBox9.ReadOnly = False
                TextBox10.ReadOnly = False
            End If
        End If
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
            func1.gccia = Label26.Text
            dt1 = dts1.consultar_detalle_guia(func1)
            If dt1.Rows.Count <> 0 Then

                DataGridView6.DataSource = dt1
                If TextBox14.Text = 1 Then
                    Dim w1, w2, num1, w3 As Integer
                    w2 = DataGridView1.Rows.Count
                    num1 = 0
                    w1 = DataGridView6.Rows.Count
                    DataGridView1.Rows.Add(w1)

                    Dim Rsr1991 As SqlDataReader
                    Dim sql1011 As String = "select nombre,codigo from TABLA_RUBROS where codigo ='" + Trim(DataGridView6.Rows(0).Cells(9).Value) + "'"
                    Dim cmd1011 As New SqlCommand(sql1011, conx)
                    Rsr1991 = cmd1011.ExecuteReader()
                    Dim l As Integer
                    l = DataGridView1.Rows.Count
                    If Rsr1991.Read() Then

                        TextBox16.Text = Trim(Rsr1991(1))
                        TextBox17.Text = Trim(Rsr1991(0))
                        TextBox13.Text = Trim(Rsr1991(0))
                        TextBox24.Text = Trim(Rsr1991(0))
                    Else
                        TextBox13.Text = ""
                    End If
                    Rsr1991.Close()

                    For i6 = w2 To w1 + w2 - 1
                        If w2 > 0 Then

                            w3 = w2 + 1
                            DataGridView1.Rows(i6).Cells(0).Value = w3

                            DataGridView1.Rows(i6).Cells(9).Value = DataGridView6.Rows(i6 - w2).Cells(1).Value
                            DataGridView1.Rows(i6).Cells(1).Value = DataGridView6.Rows(i6 - w2).Cells(2).Value
                            DataGridView1.Rows(i6).Cells(2).Value = DataGridView6.Rows(i6 - w2).Cells(3).Value
                            DataGridView1.Rows(i6).Cells(3).Value = DataGridView6.Rows(i6 - w2).Cells(7).Value
                            DataGridView1.Rows(i6).Cells(4).Value = DataGridView6.Rows(i6 - w2).Cells(6).Value

                            DataGridView1.Rows(i6).Cells(10).Value = "009"
                            DataGridView1.Rows(i6).Cells(11).Value = TextBox18.Text & "-" & TextBox19.Text
                            DataGridView1.Rows(i6).Cells(12).Value = DataGridView6.Rows(i6 - w2).Cells(8).Value

                        Else

                            num1 = num1 + 1 + w2
                            DataGridView1.Rows(i6).Cells(0).Value = num1

                            DataGridView1.Rows(i6).Cells(9).Value = DataGridView6.Rows(i6 - w2).Cells(1).Value
                            DataGridView1.Rows(i6).Cells(1).Value = DataGridView6.Rows(i6 - w2).Cells(2).Value
                            DataGridView1.Rows(i6).Cells(2).Value = DataGridView6.Rows(i6 - w2).Cells(3).Value
                            DataGridView1.Rows(i6).Cells(3).Value = DataGridView6.Rows(i6 - w2).Cells(7).Value
                            DataGridView1.Rows(i6).Cells(4).Value = DataGridView6.Rows(i6 - w2).Cells(6).Value

                            DataGridView1.Rows(i6).Cells(10).Value = "009"
                            DataGridView1.Rows(i6).Cells(11).Value = TextBox18.Text & "-" & TextBox19.Text
                            DataGridView1.Rows(i6).Cells(12).Value = DataGridView6.Rows(i6 - w2).Cells(8).Value

                        End If





                    Next

                End If


            End If
        Catch ex As Exception

            MsgBox("NO EXISTE ")

        End Try
    End Sub
    Private Sub mostrar_detalle_nsa()
        Try
            'Dim a9
            'Dim cant9, sum9 As Double
            Dim func1 As New vnia
            Dim dts1 As New fnsa

            func1.galmacen = TextBox10.Text
            Select Case TextBox8.Text.Length

                Case "1" : TextBox8.Text = "00000" & "" & TextBox8.Text
                Case "2" : TextBox8.Text = "0000" & "" & TextBox8.Text
                Case "3" : TextBox8.Text = "000" & "" & TextBox8.Text
                Case "4" : TextBox8.Text = "00" & "" & TextBox8.Text
                Case "5" : TextBox8.Text = "0" & "" & TextBox8.Text
                Case "6" : TextBox8.Text = TextBox8.Text
            End Select
            Select Case TextBox9.Text.Length

                Case "1" : TextBox9.Text = "0" & "" & TextBox9.Text
                Case "2" : TextBox9.Text = TextBox9.Text

            End Select
            func1.gncom = TextBox9.Text & "" & TextBox8.Text
            func1.gano = Label17.Text
            func1.gccia = Label26.Text
            dt8 = dts1.mostrar_detalle(func1)
            If dt8.Rows.Count <> 0 Then

                DataGridView8.DataSource = dt8
                If TextBox14.Text = 1 Then
                    Dim w1, w2, num1, w3 As Integer
                    w2 = DataGridView1.Rows.Count
                    num1 = 0
                    w1 = DataGridView8.Rows.Count
                    DataGridView1.Rows.Add(w1)

                    Dim Rsr199125 As SqlDataReader
                    Dim sql101125 As String = "select nombre,codigo from TABLA_RUBROS where codigo ='" + Trim(DataGridView8.Rows(0).Cells(9).Value) + "'"
                    Dim cmd101125 As New SqlCommand(sql101125, conx)
                    Rsr199125 = cmd101125.ExecuteReader()
                    Dim l As Integer
                    l = DataGridView1.Rows.Count
                    If Rsr199125.Read() Then

                        TextBox16.Text = Trim(Rsr199125(1))
                        TextBox17.Text = Trim(Rsr199125(0))
                        TextBox13.Text = Trim(Rsr199125(0))
                        TextBox24.Text = Trim(Rsr199125(0))
                    Else
                        TextBox13.Text = ""
                    End If
                    Rsr199125.Close()
                    For i6 = w2 To w1 + w2 - 1
                        If w2 > 0 Then

                            w3 = w2 + 1
                            DataGridView1.Rows(i6).Cells(0).Value = w3
                            DataGridView1.Rows(i6).Cells(9).Value = DataGridView8.Rows(i6 - w2).Cells(1).Value
                            DataGridView1.Rows(i6).Cells(1).Value = DataGridView8.Rows(i6 - w2).Cells(2).Value
                            DataGridView1.Rows(i6).Cells(2).Value = DataGridView8.Rows(i6 - w2).Cells(3).Value
                            DataGridView1.Rows(i6).Cells(3).Value = DataGridView8.Rows(i6 - w2).Cells(7).Value
                            DataGridView1.Rows(i6).Cells(4).Value = DataGridView8.Rows(i6 - w2).Cells(6).Value
                            DataGridView1.Rows(i6).Cells(10).Value = TextBox10.Text
                            DataGridView1.Rows(i6).Cells(11).Value = TextBox10.Text & "" & TextBox9.Text & "" & TextBox8.Text
                            DataGridView1.Rows(i6).Cells(12).Value = DataGridView8.Rows(i6 - w2).Cells(8).Value

                        Else

                            num1 = num1 + 1 + w2
                            DataGridView1.Rows(i6).Cells(0).Value = num1
                            DataGridView1.Rows(i6).Cells(9).Value = DataGridView8.Rows(i6 - w2).Cells(1).Value
                            DataGridView1.Rows(i6).Cells(1).Value = DataGridView8.Rows(i6 - w2).Cells(2).Value
                            DataGridView1.Rows(i6).Cells(2).Value = DataGridView8.Rows(i6 - w2).Cells(3).Value
                            DataGridView1.Rows(i6).Cells(3).Value = DataGridView8.Rows(i6 - w2).Cells(7).Value
                            DataGridView1.Rows(i6).Cells(4).Value = DataGridView8.Rows(i6 - w2).Cells(6).Value
                            DataGridView1.Rows(i6).Cells(10).Value = TextBox10.Text
                            DataGridView1.Rows(i6).Cells(11).Value = TextBox10.Text & "" & TextBox9.Text & "" & TextBox8.Text
                            DataGridView1.Rows(i6).Cells(12).Value = DataGridView8.Rows(i6 - w2).Cells(8).Value

                        End If





                    Next


                End If

            End If
        Catch ex As Exception

            MsgBox("NO EXISTE ")

        End Try
    End Sub




    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, a9, fila As Integer
        Dim cant10, sum10, cant9, sum9, cant8, sum8, sum7, cant7 As Double
        i = DataGridView1.Rows.Count
        fila = e.RowIndex

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cantidad" Then
            DataGridView1.Rows(fila).Cells(8).Value = DataGridView1.Rows(fila).Cells(5).Value * DataGridView1.Rows(fila).Cells(4).Value
            DataGridView1.Rows(fila).Cells(6).Value = DataGridView1.Rows(fila).Cells(8).Value / 1.18
            DataGridView1.Rows(fila).Cells(7).Value = DataGridView1.Rows(fila).Cells(8).Value - DataGridView1.Rows(fila).Cells(6).Value
            Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
            Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n2"
            Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Precio Unitario" Then
            DataGridView1.Rows(fila).Cells(8).Value = DataGridView1.Rows(fila).Cells(5).Value * DataGridView1.Rows(fila).Cells(4).Value
            DataGridView1.Rows(fila).Cells(6).Value = DataGridView1.Rows(fila).Cells(8).Value / 1.18
            DataGridView1.Rows(fila).Cells(7).Value = DataGridView1.Rows(fila).Cells(8).Value - DataGridView1.Rows(fila).Cells(6).Value
            Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n2"
            Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
            Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
        End If
        For a9 = 0 To i - 1
            cant10 = Val(DataGridView1.Rows(a9).Cells(6).Value)
            sum10 = cant10 + Val(sum10)
            cant9 = Val(DataGridView1.Rows(a9).Cells(7).Value)
            sum9 = cant9 + Val(sum9)
            cant8 = Val(DataGridView1.Rows(a9).Cells(8).Value)
            sum8 = cant8 + Val(sum8)
            cant7 = Val(DataGridView1.Rows(a9).Cells(4).Value)
            sum7 = cant7 + Val(sum7)

        Next
        TextBox29.Text = sum8.ToString("N2")
        TextBox28.Text = sum9.ToString("N2")

        TextBox26.Text = sum10.ToString("N2")
        TextBox25.Text = sum7.ToString("N2")


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim I18 As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then

            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

            I18 = DataGridView1.Rows.Count
            MsgBox(I18)
            'For i = 1 To I18
            '    MsgBox(DataGridView1.Rows(i).Cells(0).Value)
            '    MsgBox(DataGridView1.Rows(i - 1).Cells(0).Value)
            '    If DataGridView1.Rows(i).Cells(0).Value < DataGridView1.Rows(i - 1).Cells(0).Value Then
            '        num2 = DataGridView1.Rows(i - 1).Cells(0).Value + 1
            '        i17 = i

            '    End If
            'Next


            For i1 = 0 To I18 - 1

                DataGridView1.Rows(i1).Cells(0).Value = i1 + 1
            Next
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i5 As Integer
        Dim va1, va2 As String

        DataGridView1.Rows.Add(1)
        i5 = DataGridView1.Rows.Count

        If i5 = 1 Then
            va1 = 1
            Select Case va1.Length

                Case "1" : DataGridView1.Rows(i5 - 1).Cells(0).Value = "00" & "" & va1
                Case "2" : DataGridView1.Rows(i5 - 1).Cells(0).Value = "0" & "" & va1
                Case "3" : DataGridView1.Rows(i5 - 1).Cells(0).Value = va1
            End Select
            DataGridView1.Rows(i5 - 1).Cells(1).Value = "0005"
        Else
            va2 = DataGridView1.Rows(i5 - 2).Cells(0).Value + 1
            Select Case va2.Length

                Case "1" : DataGridView1.Rows(i5 - 1).Cells(0).Value = "00" & "" & va2
                Case "2" : DataGridView1.Rows(i5 - 1).Cells(0).Value = "0" & "" & va2
                Case "3" : DataGridView1.Rows(i5 - 1).Cells(0).Value = va2
            End Select
            DataGridView1.Rows(i5 - 1).Cells(1).Value = "0005"
        End If


    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

    End Sub







    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
        If CheckBox2.Checked = True Then
            Dim i8 As Integer
            i8 = DataGridView1.RowCount
            For i = 0 To i8 - 1
                DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(4).Value * DataGridView1.Rows(i).Cells(5).Value
                DataGridView1.Rows(i).Cells(6).Value = DataGridView1.Rows(i).Cells(8).Value / 1.18
                DataGridView1.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(8).Value - DataGridView1.Rows(i).Cells(6).Value
                Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"

            Next

            For a9 = 0 To i8 - 1
                cant10 = Val(DataGridView1.Rows(a9).Cells(6).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(a9).Cells(7).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(a9).Cells(8).Value)
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

                    DataGridView1.Rows(i).Cells(6).Value = DataGridView1.Rows(i).Cells(4).Value * DataGridView1.Rows(i).Cells(5).Value

                    DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(8).Value * 1.18
                    DataGridView1.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(8).Value - DataGridView1.Rows(i).Cells(6).Value
                    Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                Next

                For a9 = 0 To i8 - 1
                    cant10 = Val(DataGridView1.Rows(a9).Cells(6).Value)
                    sum10 = cant10 + Val(sum10)
                    cant9 = Val(DataGridView1.Rows(a9).Cells(7).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(DataGridView1.Rows(a9).Cells(8).Value)
                    sum8 = cant8 + Val(sum8)

                Next
                TextBox29.Text = sum8.ToString("N2")
                TextBox28.Text = sum9.ToString("N2")

                TextBox26.Text = sum10.ToString("N2")
            End If
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
        If CheckBox1.Checked = False Then

            Dim i8 As Integer
            i8 = DataGridView1.RowCount
            For i = 0 To i8 - 1
                CheckBox2.Enabled = False
                CheckBox2.Checked = False
                DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(4).Value * DataGridView1.Rows(i).Cells(5).Value
                DataGridView1.Rows(i).Cells(6).Value = DataGridView1.Rows(i).Cells(4).Value * DataGridView1.Rows(i).Cells(5).Value
                Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
                DataGridView1.Rows(i).Cells(7).Value = "0.00"

            Next

            For a9 = 0 To i8 - 1
                cant10 = Val(DataGridView1.Rows(a9).Cells(6).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(a9).Cells(7).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(a9).Cells(8).Value)
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

                    DataGridView1.Rows(i).Cells(6).Value = DataGridView1.Rows(i).Cells(4).Value * DataGridView1.Rows(i).Cells(5).Value

                    DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(8).Value * 1.18
                    DataGridView1.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(8).Value - DataGridView1.Rows(i).Cells(6).Value
                    Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                Next

                For a9 = 0 To i8 - 1
                    cant10 = Val(DataGridView1.Rows(a9).Cells(6).Value)
                    sum10 = cant10 + Val(sum10)
                    cant9 = Val(DataGridView1.Rows(a9).Cells(7).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(DataGridView1.Rows(a9).Cells(8).Value)
                    sum8 = cant8 + Val(sum8)

                Next
                TextBox29.Text = sum8.ToString("N2")
                TextBox28.Text = sum9.ToString("N2")

                TextBox26.Text = sum10.ToString("N2")
            End If

        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        Dim i, a9, fila As Integer
        Dim cant10, sum10, cant9, sum9, cant8, sum8, cant7, sum7 As Double
        i = DataGridView1.Rows.Count
        fila = e.RowIndex

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cantidad" Then
            Try
                DataGridView1.Rows(fila).Cells(8).Value = Val(DataGridView1.Rows(fila).Cells(5).Value) * Val(DataGridView1.Rows(fila).Cells(4).Value)
                DataGridView1.Rows(fila).Cells(6).Value = Val(DataGridView1.Rows(fila).Cells(8).Value) / 1.18
                DataGridView1.Rows(fila).Cells(7).Value = Val(DataGridView1.Rows(fila).Cells(8).Value) - Val(DataGridView1.Rows(fila).Cells(6).Value)
                Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
            Catch ex As Exception
                MsgBox("NO SELECCIONO UNA FILA CORRECTA")
            End Try

        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Precio Unitario" Then
            Try
                DataGridView1.Rows(fila).Cells(8).Value = Val(DataGridView1.Rows(fila).Cells(5).Value) * Val(DataGridView1.Rows(fila).Cells(4).Value)
                DataGridView1.Rows(fila).Cells(6).Value = Val(DataGridView1.Rows(fila).Cells(8).Value) / 1.18
                DataGridView1.Rows(fila).Cells(7).Value = Val(DataGridView1.Rows(fila).Cells(8).Value) - Val(DataGridView1.Rows(fila).Cells(6).Value)
                Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
            Catch ex As Exception
                MsgBox("NO SELECCIONO UNA FILA CORRECTA")
            End Try

        End If
        For a9 = 0 To i - 1
            cant10 = Val(DataGridView1.Rows(a9).Cells(6).Value)
            sum10 = cant10 + Val(sum10)
            cant9 = Val(DataGridView1.Rows(a9).Cells(7).Value)
            sum9 = cant9 + Val(sum9)
            cant8 = Val(DataGridView1.Rows(a9).Cells(8).Value)
            sum8 = cant8 + Val(sum8)
            cant7 = Val(DataGridView1.Rows(a9).Cells(4).Value)
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

    'Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
    '    Me.Close()
    'End Sub





    Private dt5 As New DataTable

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim hj As New fventasn
            Dim hj1 As New vventas_n
            Dim NU As Integer
            Select Case TextBox8.Text.Length
                Case "1" : TextBox8.Text = "00000" & "" & TextBox8.Text
                Case "2" : TextBox8.Text = "0000" & "" & TextBox8.Text
                Case "3" : TextBox8.Text = "000" & "" & TextBox8.Text
                Case "4" : TextBox8.Text = "00" & "" & TextBox8.Text
                Case "5" : TextBox8.Text = "0" & "" & TextBox8.Text
                Case "6" : TextBox8.Text = TextBox8.Text

            End Select
            hj1.gguia_nsa = TextBox10.Text & TextBox9.Text & TextBox8.Text
            hj1.gccia = Label26.Text
            NU = hj.verificar_guia(hj1)
            If NU > 0 Then
                Dim respuesta As Integer

                respuesta = MsgBox("LA NOTA SALIDA YA ESTA REGISTRADA, DESEA AGREGARLA?", vbQuestion + vbOKCancel)
                If respuesta = 1 Then '-- 2=Cancelar;1 = Aceptar
                    cabecera_nsa()
                    mostrar_detalle_nsa()
                    TextBox13.ReadOnly = False
                    TextBox15.ReadOnly = False
                    TextBox16.ReadOnly = False
                    RadioButton1.Enabled = True
                    RadioButton2.Enabled = True
                    RadioButton3.Enabled = True
                    RadioButton4.Enabled = True
                    CheckBox1.Enabled = True
                    CheckBox2.Enabled = True
                    ComboBox5.Enabled = True
                    'ComboBox6.Enabled = True
                    DataGridView1.Enabled = True
                    DataGridView1.ReadOnly = False
                    DataGridView1.Columns(0).ReadOnly = True
                    DataGridView1.Columns(4).ReadOnly = False
                    DataGridView1.Columns(5).ReadOnly = False
                    RadioButton1.Checked = True
                    CheckBox1.Checked = True
                    CheckBox2.Checked = True
                    TextBox22.ReadOnly = False
                End If
            Else
                'DataGridView1.Rows.Clear()
                cabecera_nsa()
                mostrar_detalle_nsa()
                TextBox13.ReadOnly = False
                TextBox15.ReadOnly = False
                TextBox16.ReadOnly = False
                RadioButton1.Enabled = True
                RadioButton2.Enabled = True
                RadioButton3.Enabled = True
                RadioButton4.Enabled = True
                CheckBox1.Enabled = True
                CheckBox2.Enabled = True
                ComboBox5.Enabled = True
                'ComboBox6.Enabled = True
                DataGridView1.Enabled = True
                DataGridView1.ReadOnly = False
                DataGridView1.Columns(0).ReadOnly = True
                DataGridView1.Columns(4).ReadOnly = False
                DataGridView1.Columns(5).ReadOnly = False
                RadioButton1.Checked = True
                CheckBox1.Checked = True
                CheckBox2.Checked = True
                TextBox22.ReadOnly = False
                'TextBox8.Text = ""
            End If
        End If

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox9.Text.Length
                Case "1" : TextBox9.Text = "0" & "" & TextBox9.Text
                Case "2" : TextBox9.Text = TextBox9.Text

            End Select
            TextBox8.Select()
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub
    'Sub ENVIAR_CORREO()
    '    Dim message As New MailMessage
    '    Dim smtp As New SmtpClient
    '    Dim fk As New fnopedido
    '    Dim jj As New vnapedido
    '    'Dim corre As String
    '    'jj.gvendedor = TextBox8.Text
    '    'corre = fk.buscar_correo(jj)
    '    With message
    '        .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
    '        .To.Add("hrivera@viannysac.com,asistentecobranzas@viannysac.com")
    '        .Body = "Nota Genero el Comprobante N°" & ComboBox1.Text & "-" & TextBox21.Text & "-" & TextBox5.Text
    '        .Subject = "Cliente" & TextBox7.Text & "-" & "F_Pagp" & ComboBox5.Text
    '        .Priority = System.Net.Mail.MailPriority.Normal
    '    End With

    '    With smtp
    '        .EnableSsl = True
    '        .Port = "587"
    '        .Host = "smtp.gmail.com"
    '        .Credentials = New Net.NetworkCredential("admin@viannysac.com", "$i$tema$$i$tem@$")
    '        .Send(message)
    '    End With

    '    Try
    '        MessageBox.Show("Su mensaje de correo ha sido enviado.", "Correo enviado",
    '                         MessageBoxButtons.OK)
    '    Catch ex As Exception
    '        MessageBox.Show("Error: " & ex.Message, "Error al enviar correo",
    '                         MessageBoxButtons.OK)
    '    End Try
    'End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub
    Public Function GetColumnasSize(ByVal dg As DataGridView) As Single()
        Dim values As Single() = New Single(dg.ColumnCount - 1) {}
        For i As Integer = 0 To dg.ColumnCount - 1
            values(i) = CSng(dg.Columns(i).Width)
        Next
        Return values
    End Function
    'Public Sub ExportarDatosPDF1(ByVal document As Document)
    '    ''Se crea un objeto PDFTable con el numero de columnas del DataGridView. 

    '    Dim datatable As New PdfPTable(DataGridView1.ColumnCount)
    '    'Se asignan algunas propiedades para el diseño del PDF.
    '    datatable.DefaultCell.Padding = 3
    '    Dim headerwidths As Single() = GetColumnasSize(DataGridView1)
    '    datatable.SetWidths(headerwidths)
    '    datatable.WidthPercentage = 90
    '    datatable.DefaultCell.BorderWidth = 1
    '    datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER

    '    'datatable.SetWidths(New Single() {7.0F, 5.5F, 9.5F})


    '    'datatable.TotalWidth = 500.0F
    '    'datatable.LockedWidth = True
    '    datatable.SpacingBefore = 7.0F
    '    datatable.HorizontalAlignment = Element.ALIGN_LEFT
    '    'Se crea 1el encabezado en el PDF. 
    '    Dim texto1 As New Paragraph("PROFORMA DE VENTA", New Font(Font.Name = "Tahoma", 16, Font.Bold))
    '    texto1.Alignment = Element.ALIGN_CENTER
    '    Dim texto2 As New Paragraph("CLIENTE : ---" + TextBox7.Text, New Font(Font.Name = "Tahoma", 13, Font.Bold))
    '    texto2.Alignment = Element.ALIGN_LEFT
    '    Dim texto3 As New Paragraph("CORRELATIVO : ---" + TextBox21.Text + "" + TextBox5.Text, New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    texto3.Alignment = Element.ALIGN_LEFT
    '    Dim texto7 As New Paragraph("FORMA DE PAGO : ---" + ComboBox5.Text, New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    texto7.Alignment = Element.ALIGN_LEFT
    '    Dim texto8 As New Paragraph("CONCEPTO : ---" + TextBox13.Text, New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    texto8.Alignment = Element.ALIGN_LEFT

    '    Dim texto9 As New Paragraph("MONEDA : --- SOLES", New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    texto9.Alignment = Element.ALIGN_LEFT

    '    Dim texto10 As New Paragraph("MONEDA : --- DOLARES", New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    texto10.Alignment = Element.ALIGN_LEFT


    '    Dim texto4 As New Paragraph("V. VENTA :    " + TextBox26.Text + "", New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    texto4.Alignment = Element.ALIGN_LEFT
    '    Dim texto5 As New Paragraph("IGV      :    " + TextBox28.Text + "", New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    texto5.Alignment = Element.ALIGN_LEFT
    '    Dim texto6 As New Paragraph("P. VENTA :    " + TextBox29.Text + "", New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    texto6.Alignment = Element.ALIGN_LEFT
    '    'Se capturan los nombres de las columnas del DataGridView.

    '    For i As Integer = 0 To DataGridView1.ColumnCount - 1
    '        'datatable.AddCell(DataGridView1.Columns(i).HeaderText)
    '        Dim celda As New PdfPCell(New Phrase(DataGridView1.Columns(i).HeaderText, New Font(Font.Name = "HELVETICA", 8, Font.Bold)))

    '        datatable.AddCell(celda)
    '    Next
    '    datatable.HeaderRows = 1
    '    datatable.DefaultCell.BorderWidth = 1


    '    'Se generan las columnas del DataGridView. 
    '    For i As Integer = 0 To DataGridView1.RowCount - 1
    '        For j As Integer = 0 To DataGridView1.ColumnCount - 1
    '            'compruebo que columna contien la imagen y en que columna deseo que se muestre
    '            If j = 6 Then
    '                Dim deci As Double
    '                deci = DataGridView1(j, i).Value
    '                Dim celda As New PdfPCell(New Phrase(deci.ToString("N2"), New Font(Font.Name = "Tahoma", 10)))
    '                datatable.AddCell(celda)
    '            Else
    '                If j = 7 Then
    '                    Dim deci As Double
    '                    deci = DataGridView1(j, i).Value
    '                    Dim celda As New PdfPCell(New Phrase(deci.ToString("N2"), New Font(Font.Name = "Tahoma", 9)))
    '                    datatable.AddCell(celda)
    '                Else
    '                    If j = 8 Then
    '                        Dim deci As Double
    '                        deci = DataGridView1(j, i).Value
    '                        Dim celda As New PdfPCell(New Phrase(deci.ToString("N2"), New Font(Font.Name = "Tahoma", 9)))
    '                        datatable.AddCell(celda)
    '                    Else
    '                        'datatable.AddCell(DataGridView1(j, i).Value.ToString())
    '                        Dim celda As New PdfPCell(New Phrase(DataGridView1(j, i).Value, New Font(Font.Name = "Tahoma", 9)))
    '                        datatable.AddCell(celda)
    '                    End If
    '                End If
    '            End If


    '        Next

    '    Next

    '    'Se agrega el PDFTable al documento.

    '    document.Add(texto1)
    '    document.Add(texto2)
    '    document.Add(texto3)
    '    document.Add(texto7)
    '    document.Add(texto8)
    '    If RadioButton1.Checked = True Then
    '        document.Add(texto9)
    '    Else
    '        If RadioButton2.Checked = True Then
    '            document.Add(texto10)
    '        End If
    '    End If
    '    document.Add(datatable)
    '    document.Add(texto4)
    '    document.Add(texto5)
    '    document.Add(texto6)
    'End Sub
    'Private Sub imprimir()
    '    Try
    '        'Intentar generar el documento.

    '        Dim fuente As iTextSharp.text.pdf.BaseFont
    '        Dim doc As New Document(PageSize.A4.Rotate(), 10, 10, 10, 10)
    '        'Path que guarda el reporte en el escritorio de windows (Desktop).
    '        Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\Reporteproductos.pdf"
    '        Dim file As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)


    '        PdfWriter.GetInstance(doc, file)
    '        doc.Open()

    '        fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
    '        'Seteamos el tipo de letra y el tamaño.

    '        ExportarDatosPDF1(doc)


    '        'ExportarDatosPDF2(doc)
    '        'ExportarDatosPDF3(doc)
    '        'ExportarDatosPDF4(doc)

    '        doc.Close()
    '        Process.Start(filename)
    '    Catch ex As Exception
    '        'Si el intento es fallido, mostrar MsgBox.
    '        MessageBox.Show("No se puede generar el documento PDF.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim bg As Integer
        Select Case ComboBox5.Text.ToString.Trim
            Case "CONTADO" : bg = "0"
            Case "CRED 07 DIAS" : bg = "7"
            Case "CRED 15 DIAS" : bg = "15"
            Case "CRED 30 DIAS" : bg = "30"
            Case "CRED 45 DIAS" : bg = "45"
            Case "CRED 60 DIAS" : bg = "60"
            Case "CRED 90 DIAS" : bg = "90"
            Case "CRED 120 DIAS" : bg = "120"
            Case "CONTRA ENTREGA" : bg = "0"
            Case "CRED 20 DIAS" : bg = "20"


        End Select
        DateAdd("d", bg, DateTimePicker1.Value)

        TextBox12.Text = DateAdd("d", bg, DateTimePicker1.Value)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        TextBox6.ReadOnly = False
        TextBox13.ReadOnly = False
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        RadioButton4.Enabled = True
        GroupBox5.Enabled = True
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        ComboBox5.Enabled = True
        ComboBox6.Enabled = True
        Button5.Enabled = True
        DataGridView1.Enabled = True
        DataGridView1.ReadOnly = False
        TextBox10.ReadOnly = False
        TextBox9.ReadOnly = False
        TextBox8.ReadOnly = False
        Label22.Text = 2
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Dim respuesta As DialogResult


        respuesta = MessageBox.Show("DESEA IMPRIMIR EL REGISTRO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Form_ventan.TextBox1.Text = TextBox21.Text
            Form_ventan.TextBox2.Text = TextBox5.Text
            Form_ventan.Show()
        End If
    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox9.Select()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Try
                Dim func As New fgu_comprobante
                Dim dts As New vnia


                dts.gmes = TextBox4.Text
                dts.gano = Label17.Text
                dts.gccia = Label26.Text
                Me.TextBox3.Text = func.buscar_coprobante(dts) + 1
                Select Case TextBox3.Text.Length

                    Case "1" : TextBox3.Text = "00000" & "" & TextBox3.Text
                    Case "2" : TextBox3.Text = "0000" & "" & TextBox3.Text
                    Case "3" : TextBox3.Text = "000" & "" & TextBox3.Text
                    Case "4" : TextBox3.Text = "00" & "" & TextBox3.Text
                    Case "5" : TextBox3.Text = "0" & "" & TextBox3.Text
                    Case "6" : TextBox3.Text = TextBox3.Text
                End Select
                TextBox3.Select()
            Catch ex As Exception
                MsgBox(ex.Message)


            End Try
        Else
            If e.KeyCode = Keys.F1 Then
                BUSCAR_VENTA.TextBox2.Text = Label17.Text
                BUSCAR_VENTA.TextBox3.Text = TextBox16.Text
                BUSCAR_VENTA.Label1.Text = 1
                BUSCAR_VENTA.Show()
            End If
        End If
    End Sub

    Private Sub PictureBox2_MouseHover(sender As Object, e As EventArgs)
        ToolTip1.SetToolTip(PictureBox4, "GUARDAR INFORMACION")
        ToolTip1.ToolTipTitle = "REGISTRO VENTAS"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub
    Dim Rsr2, Rsr3, Rsr4, Rsr5 As SqlDataReader

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        'abrir()
        'Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE  alias_ven ='" + ComboBox6.Text + "'"
        'Dim cmd102 As New SqlCommand(sql102, conx)
        'Rsr4 = cmd102.ExecuteReader()
        'If Rsr4.Read() = True Then
        '    TextBox14.Text = Rsr4(0)
        'End If
        'Rsr4.Close()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim ft As New fregistro_venta
        Dim jh As New vregistroventa
        Dim jh1 As New vdetalleregistro
        'Dim hj As New fventasn
        'Dim hj1 As New vventas_n
        Dim ad As Integer
        Dim suma, suma1, suma2, suma3, suma4, suma5, tcambio As Double
        'hj1.gguia_nsa = TextBox18.Text & "" & TextBox19.Text
        If TextBox34.Text = "" Then
            MsgBox("ES OBLIGATORIO INGRESAR LA NOTA DE PEDIDO")
        Else
            If TextBox12.Text = "" Or ComboBox6.Text = "---SELECCIONAR --" Then
                MsgBox("EL CAMPO FECHA VENCIMIENTO ESTA VACIO")
            Else
                'insertar cabecera
                tcambio = Convert.ToDouble(TextBox11.Text)
                jh.gcperiodo_3 = Label17.Text
                jh.gncom_3 = TextBox4.Text & "" & TextBox3.Text
                jh.gtidoc_3 = TextBox31.Text
                jh.gsfactu_3 = TextBox21.Text
                jh.gnfactu_3 = TextBox5.Text
                jh.gfcom_3 = DateTimePicker1.Value

                Select Case ComboBox5.Text.ToString.Trim
                    Case "CONTADO" : jh.gcondp_3 = "01"
                    Case "CRED 07 DIAS" : jh.gcondp_3 = "02"
                    Case "CRED 15 DIAS" : jh.gcondp_3 = "03"
                    Case "CRED 30 DIAS" : jh.gcondp_3 = "04"
                    Case "CRED 45 DIAS" : jh.gcondp_3 = "05"
                    Case "CRED 60 DIAS" : jh.gcondp_3 = "06"
                    Case "CRED 90 DIAS" : jh.gcondp_3 = "07"
                    Case "CRED 120 DIAS" : jh.gcondp_3 = "08"
                    Case "CONTRA ENTREGA" : jh.gcondp_3 = "09"
                    Case "CRED 20 DIAS" : jh.gcondp_3 = "10"
                End Select
                jh.gfich_3 = TextBox6.Text
                jh.gnomb_3 = TextBox7.Text
                If RadioButton1.Checked = True Then
                    jh.gmone_3 = 1
                Else
                    If RadioButton2.Checked = True Then
                        jh.gmone_3 = 2
                    End If
                End If

                jh.gtcam_3 = TextBox11.Text
                ad = DataGridView1.Rows.Count
                If RadioButton1.Checked = True Then
                    For i = 0 To ad - 1

                        suma = suma + Val(DataGridView1.Rows(i).Cells(6).Value)
                        suma1 = suma1 + Val(DataGridView1.Rows(i).Cells(7).Value)
                        suma2 = suma2 + Val(DataGridView1.Rows(i).Cells(8).Value)
                    Next
                    suma3 = suma / tcambio
                    suma4 = suma1 / tcambio
                    suma5 = suma2 / tcambio

                    jh.gpvta1_3 = 0.00
                    jh.gvvta1_3 = suma
                    jh.gigv1_3 = suma1
                    jh.gtot1_3 = suma2
                    jh.gpvta2_3 = 0.00
                    jh.gvvta2_3 = suma3
                    jh.gigv2_3 = suma4
                    jh.gtot2_3 = suma5
                Else
                    If RadioButton2.Checked = True Then
                        For i = 0 To ad - 1

                            suma = suma + Val(DataGridView1.Rows(i).Cells(6).Value)
                            suma1 = suma1 + Val(DataGridView1.Rows(i).Cells(7).Value)
                            suma2 = suma2 + Val(DataGridView1.Rows(i).Cells(8).Value)
                        Next
                        suma3 = suma * tcambio
                        suma4 = suma1 * tcambio
                        suma5 = suma2 * tcambio
                        jh.gpvta1_3 = 0.00
                        jh.gvvta1_3 = suma3
                        jh.gigv1_3 = suma4
                        jh.gtot1_3 = suma5
                        jh.gpvta2_3 = 0.00
                        jh.gvvta2_3 = suma
                        jh.gigv2_3 = suma1
                        jh.gtot2_3 = suma2
                    End If
                End If


                jh.ggloa_3 = TextBox13.Text
                If CheckBox1.Checked = True Then
                    jh.gaigv_3 = 1
                Else
                    jh.gaigv_3 = 0
                End If
                If CheckBox2.Checked = True Then
                    jh.giigv_3 = 1
                Else
                    jh.giigv_3 = 0
                End If

                jh.ganal1_3 = 0
                jh.gporven_3 = 0.00
                jh.gflag_3 = 1
                jh.gocompra_3 = TextBox34.Text

                abrir()
                Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + ComboBox6.Text + "'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()
                If Rsr2.Read() = True Then
                    jh.gvendedor_3 = Rsr2(0)

                End If

                Rsr2.Close()

                'Select Case ComboBox6.Text
                '    Case "GBEDON" : jh.gvendedor_3 = "0010"
                '    Case "VINCIO" : jh.gvendedor_3 = "0022"
                '    Case "DBRAVO" : jh.gvendedor_3 = "0023"
                '    Case "JSALINAS" : jh.gvendedor_3 = "0025"
                '    Case "GCUEVA" : jh.gvendedor_3 = "0029"
                '    Case "AMENDO" : jh.gvendedor_3 = "0026"
                '    Case "VGRAUS" : jh.gvendedor_3 = "0007"
                '    Case "VPIZARRO" : jh.gvendedor_3 = "0027"
                '    Case "JBALVIN" : jh.gvendedor_3 = "0028"
                '    Case "VSILVERIO" : jh.gvendedor_3 = "0005"
                '    Case "RMEDINA" : jh.gvendedor_3 = "0030"
                '    Case "WSALINAS" : jh.gvendedor_3 = "0034"
                '    Case "MGRAUS" : jh.gvendedor_3 = "0037"
                'End Select

                jh.gporcigv_3 = "0.18"
                jh.gtipo_venta = TextBox16.Text
                jh.gfechacan = DateTimePicker1.Value
                jh.gadeelanto = 0.00
                jh.gfecha_cpago = TextBox12.Text
                jh.gobservacion = TextBox22.Text
                jh.gempresa = Label26.Text
                Dim hh As New fregistro_venta
                Dim hh1 As New vdetalleregistro
                hh1.gncom_3a = TextBox4.Text & "" & TextBox3.Text
                hh1.gcperiodo_3a = Label17.Text
                hh1.gempresa = Label26.Text
                hh.eliminar_regventanegr(hh1)
                If ft.insertar_cabecera_VENTA(jh) = True Then
                    For i1 = 0 To ad - 1


                        jh1.gcperiodo_3a = Label17.Text
                        jh1.gncom_3a = TextBox4.Text & "" & TextBox3.Text
                        jh1.gitem_3a = DataGridView1.Rows(i1).Cells(0).Value
                        jh1.glinea_3a = DataGridView1.Rows(i1).Cells(1).Value
                        jh1.gproducto = DataGridView1.Rows(i1).Cells(2).Value
                        jh1.gcant_3a = DataGridView1.Rows(i1).Cells(4).Value
                        jh1.gunid_3a = DataGridView1.Rows(i1).Cells(3).Value

                        If RadioButton1.Checked = True Then
                            jh1.gpreun1_3a = Convert.ToDouble(DataGridView1.Rows(i1).Cells(5).Value)
                            jh1.gpvta1_3a = 0.00
                            jh1.gvvta1_3a = (DataGridView1.Rows(i1).Cells(6).Value)
                            jh1.gigv1_3a = (DataGridView1.Rows(i1).Cells(7).Value)
                            jh1.gtot1_3a = (DataGridView1.Rows(i1).Cells(8).Value)
                            jh1.gpreun2_3a = Convert.ToDouble(DataGridView1.Rows(i1).Cells(5).Value / TextBox11.Text)
                            jh1.gpvta2_3a = 0.00
                            jh1.gvvta2_3a = (DataGridView1.Rows(i1).Cells(6).Value / TextBox11.Text)
                            jh1.gigv2_3a = (DataGridView1.Rows(i1).Cells(7).Value / TextBox11.Text)
                            jh1.gtot2_3a = (DataGridView1.Rows(i1).Cells(8).Value / TextBox11.Text)
                        Else
                            If RadioButton2.Checked = True Then
                                jh1.gpreun1_3a = Convert.ToDouble(DataGridView1.Rows(i1).Cells(5).Value * TextBox11.Text)
                                jh1.gpvta1_3a = 0.00
                                jh1.gvvta1_3a = (DataGridView1.Rows(i1).Cells(6).Value * TextBox11.Text)
                                jh1.gigv1_3a = (DataGridView1.Rows(i1).Cells(7).Value * TextBox11.Text)
                                jh1.gtot1_3a = (DataGridView1.Rows(i1).Cells(8).Value * TextBox11.Text)
                                jh1.gpreun2_3a = Convert.ToDouble(DataGridView1.Rows(i1).Cells(5).Value)
                                jh1.gpvta2_3a = 0.00
                                jh1.gvvta2_3a = (DataGridView1.Rows(i1).Cells(6).Value)
                                jh1.gigv2_3a = (DataGridView1.Rows(i1).Cells(7).Value)
                                jh1.gtot2_3a = (DataGridView1.Rows(i1).Cells(8).Value)
                            End If

                        End If

                        jh1.gobser1_3a = ""
                        jh1.gporven_3a = 0.00
                        jh1.gordenp_3a = DataGridView1.Rows(i1).Cells(9).Value
                        jh1.ggrm_3a = DataGridView1.Rows(i1).Cells(11).Value
                        jh1.gflag_3a = 1
                        jh1.gpartida = DataGridView1.Rows(i1).Cells(12).Value
                        jh1.gempresa = Label26.Text
                        ft.insertar_detalle_VENTA(jh1)
                    Next
                    MsgBox("LA VENTA SE REGISTRO CON EXITO")
                    Dim hj2 As New flog
                    Dim hj1 As New vlog
                    hj1.gmodulo = "VENTAS_LIBRE"
                    hj1.gcuser = MDIParent1.Label3.Text
                    If Label22.Text = 1 Then
                        hj1.gaccion = 1
                    Else
                        hj1.gaccion = 2
                    End If

                    hj1.gpc = My.Computer.Name
                    hj1.gfecha = String.Format("{0:G}", DateTime.Now)
                    hj1.gccia = Label26.Text
                    hj1.gdetalle = TextBox21.Text & TextBox5.Text
                    hj2.insertar_log(hj1)
                    'vaciar informacion
                    Dim respuesta As DialogResult


                    respuesta = MessageBox.Show("DESEA IMPRIMIR EL REGISTRO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        Form_ventan.TextBox1.Text = TextBox21.Text
                        Form_ventan.TextBox2.Text = TextBox5.Text
                        Form_ventan.Show()
                    End If

                    ComboBox1.SelectedIndex = 0

                    ComboBox6.SelectedIndex = 0

                    Dim func As New ftcambio
                    Dim dts As New vtcambio
                    dts.gfecha = DateTimePicker1.Text
                    TextBox11.Text = func.mostrar_tipo_cambio(dts)
                    RadioButton1.Checked = True
                    RadioButton3.Checked = True
                    mostrar_correlativo_venta()
                    TextBox3.Select()
                    DataGridView1.DataSource = ""
                    DataGridView2.DataSource = ""
                    DataGridView3.DataSource = ""
                    DataGridView4.DataSource = ""
                    DataGridView5.DataSource = ""
                    DataGridView6.DataSource = ""
                    DataGridView7.DataSource = ""
                    DataGridView8.DataSource = ""
                    DataGridView9.DataSource = ""
                    RT.Clear()
                    TextBox6.Text = ""
                    TextBox7.Text = ""
                    TextBox13.Text = ""
                    TextBox15.Text = ""
                    TextBox18.Text = ""
                    TextBox19.Text = ""
                    TextBox10.Text = ""
                    TextBox9.Text = ""
                    TextBox8.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox28.Text = ""
                    TextBox29.Text = ""
                    TextBox22.Text = ""

                    Dim func32 As New vdocfac
                    Dim func132 As New fbolfac
                    func32.gdoc = TextBox31.Text
                    func32.gano = Label17.Text
                    func32.gccia = Label26.Text
                    TextBox5.Text = func132.mostrar_correlativo(func32) + 1

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

                End If
            End If
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Pedidos.Label3.Text = Label17.Text
        Pedidos.Label4.Text = 2
        Pedidos.Label5.Text = Label26.Text
        Pedidos.Label6.Text = TextBox27.Text
        Pedidos.Label7.Text = TextBox6.Text
        Pedidos.Show()
    End Sub

    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        ToolTip1.SetToolTip(PictureBox4, "EDITAR INFORMACION")
        ToolTip1.ToolTipTitle = "REGISTRO VENTAS"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub PictureBox4_MouseHover(sender As Object, e As EventArgs) Handles PictureBox4.MouseHover
        ToolTip1.SetToolTip(PictureBox4, "IMPRIMIR INFORMACION")
        ToolTip1.ToolTipTitle = "REGISTRO VENTAS"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub TextBox19_Layout(sender As Object, e As LayoutEventArgs) Handles TextBox19.Layout

    End Sub

    Private Sub TextBox4_DoubleClick(sender As Object, e As EventArgs) Handles TextBox4.DoubleClick
        ty2.Clear()
        ComboBox1.SelectedIndex = 0
        Dim func As New ftcambio
        Dim dts As New vtcambio
        dts.gfecha = DateTimePicker1.Text
        TextBox11.Text = func.mostrar_tipo_cambio(dts)
        RadioButton1.Checked = True
        RadioButton3.Checked = True
        mostrar_correlativo_venta()
        TextBox3.Select()
        Label22.Text = 1
        abrir()
        llenar_combo_box2()
    End Sub
End Class