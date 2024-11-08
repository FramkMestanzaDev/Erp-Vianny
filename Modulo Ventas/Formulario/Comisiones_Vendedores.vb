
Imports System.Data.SqlClient
Public Class Comisiones_Vendedores
    Dim DT, DT2 As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim ty, ty2, ty3 As New DataTable
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
    Dim Rsr22, Rsr222 As SqlDataReader
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        DataGridView1.DataSource = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        Dim va1, va2 As String

        va1 = ""
        va2 = ""
        If Trim(ComboBox1.Text) = "LGONZALEZ" Then
            TextBox7.Text = "3,000,000.00"
        End If
        Select Case ComboBox2.Text
            Case "ENERO" : va1 = "01"
            Case "FEBRERO" : va1 = "02"
            Case "MARZO" : va1 = "03"
            Case "ABRIL" : va1 = "04"
            Case "MAYO" : va1 = "05"
            Case "JUNIO" : va1 = "06"
            Case "JULIO" : va1 = "07"
            Case "AGOSTO" : va1 = "08"
            Case "SETIEMBRE" : va1 = "09"
            Case "OCTUBRE" : va1 = "10"
            Case "NOVIEMBRE" : va1 = "11"
            Case "DICIEMBRE" : va1 = "12"
        End Select
        'Select Case ComboBox1.Text
        '    Case "GBEDON" : va2 = "0010"
        '    Case "VINCIO" : va2 = "0022"
        '    Case "DBRAVO" : va2 = "0023"
        '    Case "JSALINAS" : va2 = "0025"
        '    Case "GCUEVA" : va2 = "0029"
        '    Case "AMENDO" : va2 = "0026"
        '    Case "VGRAUS" : va2 = "0007"
        '    Case "VPIZARRO" : va2 = "0027"
        '    Case "GBALVIN" : va2 = "0028"
        '    Case "VSILVERIO" : va2 = "0005"
        '    Case "WSALINAS" : va2 = "0034"
        'End Select
        va2 = ComboBox1.SelectedValue.ToString

        abrir()
        Dim Rs As SqlDataReader
        Dim sql As String = ("select count(id) from comsion_vendedor where id ='" + Trim(Label8.Text) + Trim(Label12.Text) + va1 + va2 + "'")
        Dim cmd As New SqlCommand(sql, conx)
        Rs = cmd.ExecuteReader()
        Rs.Read()
        If Rs(0) > 0 Then
            Dim respuesta As DialogResult

            respuesta = MessageBox.Show("YA EXISTE UN REGISTRO DE COMISIONES DEL VENDEDOR EN EL MES SOLICITADO, DESEA CARGAR LA INFORMACION ACTUALIZANDO SUS NUEVAS VENTAS?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                MsgBox("GRACIAS ESTA ES LA ACTUALIZACION")
                CheckBox1.Enabled = False
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
                Dim fgw, PP As New vfactura
                Dim fgw1 As New ffactura
                Dim suma, suma1, VA, JU, JU1 As Double
                Dim va12, BO, NP, va13 As Double
                'Select Case ComboBox1.Text
                '    Case "GBEDON" : fgw.gvendedor = "0010"
                '    Case "VINCIO" : fgw.gvendedor = "0022"
                '    Case "DBRAVO" : fgw.gvendedor = "0023"
                '    Case "JSALINAS" : fgw.gvendedor = "0025"
                '    Case "GCUEVA" : fgw.gvendedor = "0029"
                '    Case "AMENDO" : fgw.gvendedor = "0026"
                '    Case "VGRAUS" : fgw.gvendedor = "0007"
                '    Case "VPIZARRO" : fgw.gvendedor = "0027"
                '    Case "GBALVIN" : fgw.gvendedor = "0028"
                '    Case "VSILVERIO" : fgw.gvendedor = "0005"
                '    Case "WSALINAS" : fgw.gvendedor = "0034"
                'End Select
                fgw.gvendedor = ComboBox1.SelectedValue.ToString

                Select Case ComboBox2.Text
                    Case "ENERO" : fgw.gmes = "01"
                    Case "FEBRERO" : fgw.gmes = "02"
                    Case "MARZO" : fgw.gmes = "03"
                    Case "ABRIL" : fgw.gmes = "04"
                    Case "MAYO" : fgw.gmes = "05"
                    Case "JUNIO" : fgw.gmes = "06"
                    Case "JULIO" : fgw.gmes = "07"
                    Case "AGOSTO" : fgw.gmes = "08"
                    Case "SETIEMBRE" : fgw.gmes = "09"
                    Case "OCTUBRE" : fgw.gmes = "10"
                    Case "NOVIEMBRE" : fgw.gmes = "11"
                    Case "DICIEMBRE" : fgw.gmes = "12"
                End Select
                PP.gmes = fgw.gmes
                Label5.Text = ComboBox1.Text
                Label15.Text = 1
                fgw.gperiodo = Label8.Text
                PP.gperiodo = Label8.Text
                fgw.gccia = Label12.Text
                PP.gccia = Label12.Text

                If Trim(ComboBox1.Text) = "LGONZALEZ" Then
                    DT2 = fgw1.comisiones_edicion_total(PP)
                Else
                    DT2 = fgw1.comisiones_vendedor_edicion(fgw)
                End If

                DataGridView1.DataSource = DT2
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(15).Visible = False
                DataGridView1.Columns(16).Visible = False
                DataGridView1.Columns(17).Visible = False
                DataGridView1.Columns(18).Visible = False
                DataGridView1.Columns(21).Visible = False
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(2).Width = 140
                DataGridView1.Columns(3).Width = 120
                DataGridView1.Columns(4).Width = 90
                DataGridView1.Columns(6).Width = 60
                DataGridView1.Columns(7).Width = 90
                DataGridView1.Columns(12).Width = 80
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(10).DefaultCellStyle.Format = "N2"
                DataGridView1.Columns(12).DefaultCellStyle.Format = "N2"
                TextBox5.Text = DataGridView1.Rows(0).Cells(15).Value
                Dim pp1 As Double
                pp1 = (Convert.ToDouble(TextBox7.Text) - Convert.ToDouble(DataGridView1.Rows(0).Cells(15).Value)) / 5000
                TextBox6.Text = Convert.ToInt16(pp1)
                Dim ui As Integer
                ui = DataGridView1.Rows.Count

                For i = 0 To ui - 1
                    'va12 = Convert.ToDouble(DataGridView1.Rows(i).Cells(10).Value).ToString("N2")

                    'DataGridView1.Rows(i).Cells(10).Value = va12
                    suma1 = suma1 + DataGridView1.Rows(i).Cells(9).Value
                Next
                TextBox2.Text = suma1.ToString("N2")
                va13 = DataGridView1.Rows(0).Cells(15).Value
                JU1 = Convert.ToDouble(TextBox2.Text) - va13
                TextBox4.Text = JU1.ToString("N2")
                If Convert.ToDouble(TextBox2.Text) >= va13 Then
                    TextBox3.Visible = True
                    TextBox3.Text = "SI COMISIONARA PORQUE ES EN VENDEDOR LIBRE"
                Else
                    TextBox3.Visible = True
                    TextBox3.Text = "LA VENTAS NO SUPERAN LA CUOTA MENSUAL,POR LO TANTO SOLO SE LE PAGARA SU BASICO"
                    TextBox1.Text = "0.00"
                    TextBox4.Text = "0.00"
                End If
                ui = DataGridView1.Rows.Count
                For i = 0 To ui - 1
                    'If Label5.Text = "AMENDO" Or Label12.Text = "02" Or fg.gvendedor = "0027" Or fg.gvendedor = "0028" Then

                    '    DataGridView1.Rows(i).Cells(10).Value = DataGridView1.Rows(i).Cells(9).Value
                    'Else
                    BO = ((TextBox4.Text / TextBox2.Text) * 100)
                    NP = (Convert.ToDouble(DataGridView1.Rows(i).Cells(9).Value * BO)) / 100
                    DataGridView1.Rows(i).Cells(10).Value = NP
                    'End If
                    VA = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                    DataGridView1.Rows(i).Cells(12).Value = VA
                    suma = suma + DataGridView1.Rows(i).Cells(12).Value
                Next
                TextBox1.Text = suma.ToString("N2")
                Button4.Enabled = True
            End If
            Select Case ComboBox2.Text
                Case "ENERO" : va1 = "01"
                Case "FEBRERO" : va1 = "02"
                Case "MARZO" : va1 = "03"
                Case "ABRIL" : va1 = "04"
                Case "MAYO" : va1 = "05"
                Case "JUNIO" : va1 = "06"
                Case "JULIO" : va1 = "07"
                Case "AGOSTO" : va1 = "08"
                Case "SETIEMBRE" : va1 = "09"
                Case "OCTUBRE" : va1 = "10"
                Case "NOVIEMBRE" : va1 = "11"
                Case "DICIEMBRE" : va1 = "12"
            End Select
            Rs.Close()
            va2 = ComboBox1.SelectedValue.ToString
            Dim sql1022 As String = "select distinct apgg from comsion_vendedor  where id ='" + Trim(Label8.Text) + Trim(Label12.Text) + va1 + va2 + "'"
            Dim cmd1022 As New SqlCommand(sql1022, conx)
            Rsr22 = cmd1022.ExecuteReader()
            If Rsr22.Read() Then
                If Rsr22(0) = "1" Then
                    CheckBox2.Checked = True
                    Button5.Enabled = False
                Else
                    CheckBox2.Checked = False
                    Button5.Enabled = True
                End If
            End If

            Rsr22.Close()
        Else
            If Trim(TextBox5.Text) = "" Then
                MsgBox("FALTA INGRESAR EL MONTO MINIMO DE VENTAS")
            Else

                Dim valor As Double
                valor = Convert.ToDouble(TextBox7.Text) - (5000 * Convert.ToDouble(TextBox6.Text))
                TextBox5.Text = valor.ToString("N2")
                Dim respuesta As DialogResult

                respuesta = MessageBox.Show("EL MONTO MINIMO A COMISIONAR ES : " + TextBox5.Text + " ESTA SEGURO DE CONTINUAR ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)


                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim fg, HH As New vfactura
                    Dim fg1 As New ffactura

                    If ComboBox1.Text = " " Or ComboBox2.Text = " " Then
                        MsgBox("FALTA SELECCIONAR UN VENDEDOR O MES")
                    Else
                        ComboBox1.Enabled = False
                        ComboBox2.Enabled = False
                        'Select Case ComboBox1.Text
                        '    Case "GBEDON" : fg.gvendedor = "0010"
                        '    Case "VINCIO" : fg.gvendedor = "0022"
                        '    Case "DBRAVO" : fg.gvendedor = "0023"
                        '    Case "JSALINAS" : fg.gvendedor = "0025"
                        '    Case "GCUEVA" : fg.gvendedor = "0029"
                        '    Case "AMENDO" : fg.gvendedor = "0026"
                        '    Case "VGRAUS" : fg.gvendedor = "0007"
                        '    Case "VPIZARRO" : fg.gvendedor = "0027"
                        '    Case "GBALVIN" : fg.gvendedor = "0028"
                        '    Case "VSILVERIO" : fg.gvendedor = "0005"
                        '    Case "WSALINAS" : fg.gvendedor = "0034"
                        'End Select
                        fg.gvendedor = ComboBox1.SelectedValue.ToString
                        Select Case ComboBox2.Text
                            Case "ENERO" : fg.gmes = "01"
                            Case "FEBRERO" : fg.gmes = "02"
                            Case "MARZO" : fg.gmes = "03"
                            Case "ABRIL" : fg.gmes = "04"
                            Case "MAYO" : fg.gmes = "05"
                            Case "JUNIO" : fg.gmes = "06"
                            Case "JULIO" : fg.gmes = "07"
                            Case "AGOSTO" : fg.gmes = "08"
                            Case "SETIEMBRE" : fg.gmes = "09"
                            Case "OCTUBRE" : fg.gmes = "10"
                            Case "NOVIEMBRE" : fg.gmes = "11"
                            Case "DICIEMBRE" : fg.gmes = "12"
                        End Select

                        Label5.Text = ComboBox1.Text
                        Label15.Text = 2
                        fg.gperiodo = Label8.Text
                        fg.gccia = Label12.Text
                        HH.gmes = fg.gmes
                        HH.gperiodo = Label8.Text
                        HH.gccia = Label12.Text
                        If Trim(ComboBox1.Text) = "LGONZALEZ" Then

                            DT = fg1.comisiones_vendedor_detallado(HH)

                        Else
                            DT = fg1.comisiones_vendedor(fg)
                        End If

                        DataGridView1.DataSource = DT
                        Dim ui As Integer
                        Dim suma, suma1, VA, JU, JU1 As Double
                        ui = DataGridView1.Rows.Count
                        Dim BO, NP, va13 As Double
                        For i = 0 To ui - 1

                            suma1 = suma1 + DataGridView1.Rows(i).Cells(9).Value
                        Next
                        DataGridView1.Columns(13).DefaultCellStyle.BackColor = Color.DarkKhaki
                        DataGridView1.Columns(13).DefaultCellStyle.ForeColor = Color.White
                        DataGridView1.Columns(15).ReadOnly = True
                        DataGridView1.Columns(0).Visible = False
                        DataGridView1.Columns(17).Visible = False
                        DataGridView1.Columns(1).Width = 90
                        DataGridView1.Columns(2).Width = 140
                        DataGridView1.Columns(3).Width = 120
                        DataGridView1.Columns(4).Width = 90
                        DataGridView1.Columns(6).Width = 60
                        DataGridView1.Columns(7).Width = 90
                        DataGridView1.Columns(12).Width = 80
                        DataGridView1.Columns(4).Frozen = True
                        DataGridView1.Columns(1).Frozen = True
                        DataGridView1.Columns(2).Frozen = True
                        DataGridView1.Columns(3).Frozen = True
                        DataGridView1.Columns(15).DefaultCellStyle.Format = "yyyy/MM/dd"
                        TextBox2.Text = suma1.ToString("N2")
                        Button4.Enabled = True
                        'va12 = TextBox5.Text
                        va13 = TextBox5.Text
                        'JU = Convert.ToDouble(TextBox2.Text) - va12
                        JU1 = Convert.ToDouble(TextBox2.Text) - va13
                        TextBox4.Text = JU1.ToString("N2")
                        'If Label5.Text = "AMENDO" Or Label12.Text = "02" Or fg.gvendedor = "0027" Or fg.gvendedor = "0028" Then
                        '    TextBox4.Text = TextBox2.Text
                        'Else
                        '    If fg.gvendedor = "0010" Or fg.gvendedor = "0023" Or fg.gvendedor = "0034" Then
                        '        TextBox4.Text = JU1.ToString("N2")
                        '        'Label15.Text = va13
                        '    Else
                        '        TextBox4.Text = JU.ToString("N2")
                        '        'Label15.Text = va12
                        '    End If

                        'End If
                        If Convert.ToDouble(TextBox2.Text) >= va13 Then
                            TextBox3.Visible = True
                            TextBox3.Text = "SI COMISIONARA PORQUE ES EN VENDEDOR LIBRE"
                        Else
                            TextBox3.Visible = True
                            TextBox3.Text = "LA VENTAS NO SUPERAN LA CUOTA MENSUAL,POR LO TANTO SOLO SE LE PAGARA SU BASICO"
                            TextBox1.Text = "0.00"
                            TextBox4.Text = "0.00"
                        End If
                        '    If fg.gvendedor = "0010" Or fg.gvendedor = "0023" Or fg.gvendedor = "0034" Then
                        '    If Convert.ToDouble(TextBox2.Text) <= va13 Then
                        '        If Label5.Text = "AMENDO" Or Label12.Text = "02" Then
                        '            TextBox3.Visible = True
                        '            TextBox3.Text = "SI COMISIONARA PORQUE ES EN VENDEDOR LIBRE"
                        '        Else
                        '            TextBox3.Visible = True
                        '            TextBox3.Text = "LA VENTAS NO SUPERAN LA CUOTA MENSUAL,POR LO TANTO SOLO SE LE PAGARA SU BASICO"
                        '            TextBox1.Text = "0.00"
                        '        End If

                        '    Else
                        '        If Convert.ToDouble(TextBox2.Text) >= va13 Then
                        '            TextBox3.Visible = True
                        '            TextBox3.Text = "LA VENTAS  SUPERAN LA CUOTA MENSUAL, SI COMISIONARA"
                        '        End If
                        '    End If


                        'Else
                        '    If Convert.ToDouble(TextBox2.Text) <= va12 Then
                        '        If Label5.Text = "AMENDO" Or Label12.Text = "02" Or fg.gvendedor = "0027" Or fg.gvendedor = "0028" Then
                        '            TextBox3.Visible = True
                        '            TextBox3.Text = "SI COMISIONARA "
                        '        Else
                        '            TextBox3.Visible = True
                        '            TextBox3.Text = "LA VENTAS NO SUPERAN LA CUOTA MENSUAL,POR LO TANTO SOLO SE LE PAGARA SU BASICO"
                        '            TextBox1.Text = "0.00"
                        '        End If

                        '    Else
                        '        If Convert.ToDouble(TextBox2.Text) >= va12 Then
                        '            TextBox3.Visible = True
                        '            TextBox3.Text = "LA VENTAS  SUPERAN LA CUOTA MENSUAL, SI COMISIONARA"
                        '        End If
                        '    End If
                        'End If

                        ui = DataGridView1.Rows.Count
                        For i = 0 To ui - 1
                            'If Label5.Text = "AMENDO" Or Label12.Text = "02" Or fg.gvendedor = "0027" Or fg.gvendedor = "0028" Then

                            '    DataGridView1.Rows(i).Cells(10).Value = DataGridView1.Rows(i).Cells(9).Value
                            'Else
                            BO = ((TextBox4.Text / TextBox2.Text) * 100)
                            NP = (DataGridView1.Rows(i).Cells(9).Value * BO) / 100
                            DataGridView1.Rows(i).Cells(10).Value = NP.ToString("N2")
                            'End If


                            VA = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                            DataGridView1.Rows(i).Cells(12).Value = VA.ToString("N2")
                            suma = suma + DataGridView1.Rows(i).Cells(12).Value

                            'Dim y As String
                            'y = Format(DataGridView1.Rows(i).Cells(15).Value, “yyyy/MM/dd”).ToString
                            'DataGridView1.Rows(i).Cells(15).Value = y

                        Next
                        TextBox1.Text = suma


                        If Trim(ComboBox1.Text) = "GBALVIN" Or Trim(ComboBox1.Text) = "VSILVERIO" Then
                            Dim lp As Integer
                            lp = DataGridView1.Rows.Count
                            If (Convert.ToDouble(TextBox2.Text) >= 100000.0) And (Convert.ToDouble(TextBox2.Text) <= 300000.0) Then
                                Dim tot, SUMA22 As Double
                                For i = 0 To lp - 1
                                    If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                        DataGridView1.Rows(i).Cells(11).Value = 0.4
                                        tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                        DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                        SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                    End If

                                Next
                                TextBox1.Text = SUMA22
                            Else

                                If (Convert.ToDouble(TextBox2.Text) >= 301001.0) And (Convert.ToDouble(TextBox2.Text) <= 450000.0) Then
                                    Dim tot, SUMA22 As Double
                                    For i = 0 To lp - 1
                                        If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                            DataGridView1.Rows(i).Cells(11).Value = 0.45
                                            tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                            DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                            SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                        End If

                                    Next
                                    TextBox1.Text = SUMA22
                                Else
                                    If (Convert.ToDouble(TextBox2.Text) >= 450001.0) And (Convert.ToDouble(TextBox2.Text) <= 600000.0) Then
                                        Dim tot, SUMA22 As Double
                                        For i = 0 To lp - 1
                                            If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                                DataGridView1.Rows(i).Cells(11).Value = 0.5
                                                tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                                DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                                SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                            End If

                                        Next
                                        TextBox1.Text = SUMA22
                                    Else
                                        If (Convert.ToDouble(TextBox2.Text) >= 600001.0) And (Convert.ToDouble(TextBox2.Text) <= 750000.0) Then
                                            Dim tot, SUMA22 As Double
                                            For i = 0 To lp - 1
                                                If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                                    DataGridView1.Rows(i).Cells(11).Value = 0.55
                                                    tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                                    DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                                    SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                                End If

                                            Next
                                            TextBox1.Text = SUMA22
                                        Else
                                            If (Convert.ToDouble(TextBox2.Text) >= 750001.0) And (Convert.ToDouble(TextBox2.Text) <= 900000.0) Then
                                                Dim tot, SUMA22 As Double
                                                For i = 0 To lp - 1
                                                    If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                                        DataGridView1.Rows(i).Cells(11).Value = 0.6
                                                        tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                                        DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                                        SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                                    End If

                                                Next
                                                TextBox1.Text = SUMA22
                                            Else
                                                If (Convert.ToDouble(TextBox2.Text) >= 900001.0) And (Convert.ToDouble(TextBox2.Text) <= 1050000.0) Then
                                                    Dim tot, SUMA22 As Double
                                                    For i = 0 To lp - 1
                                                        If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                                            DataGridView1.Rows(i).Cells(11).Value = 0.65
                                                            tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                                            DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                                            SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                                        End If

                                                    Next
                                                    TextBox1.Text = SUMA22
                                                Else
                                                    If (Convert.ToDouble(TextBox2.Text) >= 1050001.0) And (Convert.ToDouble(TextBox2.Text) <= 1200000.0) Then
                                                        Dim tot, SUMA22 As Double
                                                        For i = 0 To lp - 1
                                                            If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                                                DataGridView1.Rows(i).Cells(11).Value = 0.7
                                                                tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                                                DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                                                SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                                            End If

                                                        Next
                                                        TextBox1.Text = SUMA22
                                                    Else
                                                        If (Convert.ToDouble(TextBox2.Text) >= 1200001.0) And (Convert.ToDouble(TextBox2.Text) <= 1350000.0) Then
                                                            Dim tot, SUMA22 As Double
                                                            For i = 0 To lp - 1
                                                                If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                                                    DataGridView1.Rows(i).Cells(11).Value = 0.75
                                                                    tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                                                    DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                                                    SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                                                End If

                                                            Next
                                                            TextBox1.Text = SUMA22
                                                        Else
                                                            If (Convert.ToDouble(TextBox2.Text) >= 1350001.0) And (Convert.ToDouble(TextBox2.Text) <= 1500000.0) Then
                                                                Dim tot, SUMA22 As Double
                                                                For i = 0 To lp - 1
                                                                    If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                                                        DataGridView1.Rows(i).Cells(11).Value = 0.8
                                                                        tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                                                        DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                                                        SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                                                    End If

                                                                Next
                                                                TextBox1.Text = SUMA22
                                                            Else
                                                                If (Convert.ToDouble(TextBox2.Text) >= 1500001.0) And (Convert.ToDouble(TextBox2.Text) <= 1650000.0) Then
                                                                    Dim tot, SUMA22 As Double
                                                                    For i = 0 To lp - 1
                                                                        If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                                                            DataGridView1.Rows(i).Cells(11).Value = 0.85
                                                                            tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                                                            DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                                                            SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                                                        End If

                                                                    Next
                                                                    TextBox1.Text = SUMA22
                                                                Else
                                                                    If (Convert.ToDouble(TextBox2.Text) >= 1650001.0) And (Convert.ToDouble(TextBox2.Text) <= 1800000.0) Then
                                                                        Dim tot, SUMA22 As Double
                                                                        For i = 0 To lp - 1
                                                                            If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                                                                DataGridView1.Rows(i).Cells(11).Value = 0.9
                                                                                tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                                                                DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                                                                SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                                                            End If

                                                                        Next
                                                                        TextBox1.Text = SUMA22
                                                                    Else
                                                                        If (Convert.ToDouble(TextBox2.Text) >= 1800001.0) And (Convert.ToDouble(TextBox2.Text) <= 1950000.0) Then
                                                                            Dim tot, SUMA22 As Double
                                                                            For i = 0 To lp - 1
                                                                                If Trim(DataGridView1.Rows(i).Cells(3).Value) = "MANUFACTURA" Then
                                                                                    DataGridView1.Rows(i).Cells(11).Value = 0.95
                                                                                    tot = (DataGridView1.Rows(i).Cells(10).Value * DataGridView1.Rows(i).Cells(11).Value) / 100
                                                                                    DataGridView1.Rows(i).Cells(12).Value = tot.ToString("N2")
                                                                                    SUMA22 = SUMA22 + DataGridView1.Rows(i).Cells(12).Value
                                                                                End If

                                                                            Next
                                                                            TextBox1.Text = SUMA22
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
                    Button5.Enabled = True
                End If
            End If
        End If
        Rs.Close()



    End Sub



    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        Button1.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim va11, va21 As String
        va11 = ""
        va21 = ""
        If Label15.Text = 0 Then
            'Select Case TextBox3.Text
            '    Case "GBEDON" : va11 = "0010"
            '    Case "VINCIO" : va11 = "0022"
            '    Case "DBRAVO" : va11 = "0023"
            '    Case "JSALINAS" : va11 = "0025"
            '    Case "GCUEVA" : va11 = "0029"
            '    Case "AMENDO" : va11 = "0026"
            '    Case "VGRAUS" : va11 = "0007"
            '    Case "VPIZARRO" : va11 = "0027"
            '    Case "GBALVIN" : va11 = "0028"
            '    Case "VSILVERIO" : va11 = "0005"
            '    Case "WSALINAS" : va11 = "0034"
            'End Select

            va11 = ComboBox1.SelectedValue.ToString
        Else
            If Label15.Text = 1 Then
                'Select Case ComboBox1.Text
                '    Case "GBEDON" : va11 = "0010"
                '    Case "VINCIO" : va11 = "0022"
                '    Case "DBRAVO" : va11 = "0023"
                '    Case "JSALINAS" : va11 = "0025"
                '    Case "GCUEVA" : va11 = "0029"
                '    Case "AMENDO" : va11 = "0026"
                '    Case "VGRAUS" : va11 = "0007"
                '    Case "VPIZARRO" : va11 = "0027"
                '    Case "GBALVIN" : va11 = "0028"
                '    Case "VSILVERIO" : va11 = "0005"
                '    Case "WSALINAS" : va11 = "0034"
                'End Select
                va11 = ComboBox1.SelectedValue.ToString
            End If

        End If


        Select Case ComboBox2.Text
            Case "ENERO" : va21 = "01"
            Case "FEBRERO" : va21 = "02"
            Case "MARZO" : va21 = "03"
            Case "ABRIL" : va21 = "04"
            Case "MAYO" : va21 = "05"
            Case "JUNIO" : va21 = "06"
            Case "JULIO" : va21 = "07"
            Case "AGOSTO" : va21 = "08"
            Case "SETIEMBRE" : va21 = "09"
            Case "OCTUBRE" : va21 = "10"
            Case "NOVIEMBRE" : va21 = "11"
            Case "DICIEMBRE" : va21 = "12"
        End Select

        Dim sql1022 As String = "select distinct apgg from comsion_vendedor  where id ='" + Trim(Label8.Text) + Trim(Label12.Text) + va21 + va11 + "'"
        Dim cmd1022 As New SqlCommand(sql1022, conx)
        Rsr222 = cmd1022.ExecuteReader()
        If Rsr222.Read() Then
            If Rsr222(0) = "1" Then
                Reporte_comisiones.TextBox1.Text = Trim(Label8.Text) + Trim(Label12.Text) + va21 + va11
                Reporte_comisiones.Show()
            Else
                MsgBox("LA COMISION NO ESTA APROBADA POR LA GERENCIA, POR LO QUE NO SE PUEDE IMPRIMIR")
            End If
        End If

        Rsr222.Close()


    End Sub
    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("sELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  rpm_ven='2' and ccia_ven ='01' and admin_ven ='0'", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "alias_ven"
            ComboBox1.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box3()
        Try

            conn = New SqlDataAdapter("sELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  rpm_ven='2' and ccia_ven ='01' and admin_ven in (2,3)", conx)
            conn.Fill(ty3)
            ComboBox1.DataSource = ty3
            ComboBox1.DisplayMember = "alias_ven"
            ComboBox1.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.Text = "-- SELECCIONAR --" Or ComboBox2.Text = "-- SELECCIONAR --" Then
            MsgBox("FALTA SELECCIONAR UN VENDEDOR O MES")
        Else
            Dim jk As New Exportar
            jk.llenarExcel(DataGridView1)
        End If
    End Sub

    Private Sub Comisiones_Vendedores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox2.SelectedIndex = DateTime.Now.ToString("MM") - 1

        abrir()
        llenar_combo_box2()
        If MDIParent1.Label4.Text = "ADMINISTRADOR" Then
            CheckBox2.Visible = True
        Else
            CheckBox2.Visible = False
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim fg As New vfactura
        Dim fg1 As New ffactura
        Dim oo As Integer
        Dim JJ As String
        oo = DataGridView1.Rows.Count
        For i = 0 To oo - 1

            If Trim(DataGridView1.Rows(i).Cells(1).Value) = "INGRESOS" Then
                If Mid(DataGridView1.Rows(i).Cells(6).Value, 1, 1) = "F" Then
                    fg.gdoc = "001"
                Else
                    If Mid(DataGridView1.Rows(i).Cells(6).Value, 1, 1) = "B" Then
                        fg.gdoc = "003"
                    End If

                End If

                fg.gndoc = DataGridView1.Rows(i).Cells(6).Value + "-" + DataGridView1.Rows(i).Cells(7).Value
                fg.gCLIENTE = DataGridView1.Rows(i).Cells(2).Value
                fg.gperiodo = Label8.Text
                fg.gccia = Label12.Text
                JJ = fg1.mostrar_estado_facbol2(fg)
                If JJ = "False" Then
                    DataGridView1.Rows(i).Cells(13).Value = fg1.mostrar_estado_facbol(fg)
                Else
                    DataGridView1.Rows(i).Cells(13).Value = fg1.mostrar_estado_facbol2(fg)
                End If

            Else
                DataGridView1.Rows(i).Cells(13).Value = "CANCELADO"
            End If

        Next
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, fila As Integer
        Dim GH, suma As Double
        Dim U As Integer
        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        suma = 0
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PORCENTAJE" Then
            Try

                GH = (DataGridView1.Rows(fila).Cells(11).Value * DataGridView1.Rows(fila).Cells(10).Value) / 100
                DataGridView1.Rows(fila).Cells(12).Value = GH

                For U = 0 To i - 1
                    suma = suma + DataGridView1.Rows(U).Cells(12).Value
                Next

            Catch ex As Exception
                MsgBox("NO SELECCIONO UNA FILA CORRECTA")
            End Try

        End If

        TextBox1.Text = suma.ToString("N2")
    End Sub

    Public Sub NumConFrac(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar = "." And Not CajaTexto.Text.IndexOf(".") Then
            e.Handled = True
        ElseIf e.KeyChar = "." Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Public Sub NumConFrac2(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False

        Else
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox5_LostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus
        Dim NumAuxiliar As Double
        NumAuxiliar = TextBox5.Text

        TextBox5.Text = FormatNumber(NumAuxiliar, 2)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim op As Integer
        op = DataGridView1.Rows.Count
        Dim suma As Double
        For i = 0 To op - 1
            'MsgBox(DataGridView1.Rows(i).Cells(8).Value)
            'MsgBox(DataGridView1.Rows(i).Cells(11).Value)
            'MsgBox(DataGridView1.Rows(i).Cells(12).Value)
            'MsgBox(DataGridView1.Rows(i).Cells(14).Value)
            Dim ko As Date
            ko = DataGridView1.Rows(i).Cells(15).Value
            'MsgBox(Replace(ko.ToString("yyyy-MM-dd"), "-", ""))
            Dim l As String

            Dim Rsr1991 As SqlDataReader
            Dim sql1011 As String = "select distinct top 1  correl_tabla from tabla_comisiones where fecha <= '" + Replace(ko.ToString("yyyy-MM-dd"), "-", "") + "' order by correl_tabla desc"
            Dim cmd1011 As New SqlCommand(sql1011, conx)
            Rsr1991 = cmd1011.ExecuteReader()
            If Rsr1991.Read() Then
                'MsgBox(DataGridView1.Rows(i).Cells(15).Value)
                'MsgBox(Rsr1991(0))
                l = Rsr1991(0)
                Rsr1991.Close()


            Else
                l = 0
            End If
            Rsr1991.Close()
            Dim precio As String
            precio = Convert.ToString(DataGridView1.Rows(i).Cells(8).Value)

            Dim Rsr19912 As SqlDataReader
            Dim sql10112 As String = "select  case when fpagocontado = '01' then  CASE WHEN preciocontado >'" + precio + "' THEN operacioncontado ELSE comisioncontado END else '0.5'  end AS COMISION from tabla_comisiones where correl_tabla ='" + l + "  ' and almacen ='" + Trim(DataGridView1.Rows(i).Cells(14).Value) + "' AND rubro_codigo ='" + Trim(DataGridView1.Rows(i).Cells(17).Value) + "'"
            Dim cmd10112 As New SqlCommand(sql10112, conx)
            Rsr19912 = cmd10112.ExecuteReader()
            If Rsr19912.Read() Then
                DataGridView1.Rows(i).Cells(11).Value = Rsr19912(0)
                DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(10).Value * Rsr19912(0)
            End If
            Rsr19912.Close()

            suma = suma + DataGridView1.Rows(i).Cells(12).Value
        Next
        TextBox1.Text = suma
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        NumConFrac(TextBox5, e)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        DataGridView1.DataSource = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = "0.00"
        Label5.Text = ""
        TextBox3.Visible = False
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        CheckBox1.Enabled = True
        TextBox7.Text = "150,000.00"
        DT2.Clear()
        DT.Clear()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ty3.Clear()
            abrir()
            llenar_combo_box3()
        Else
            ty2.Clear()
            abrir()
            llenar_combo_box2()
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim y As Integer
        Dim va1, va2 As String
        va1 = ""
        va2 = ""
        y = DataGridView1.Rows.Count
        Select Case ComboBox2.Text
            Case "ENERO" : va1 = "01"
            Case "FEBRERO" : va1 = "02"
            Case "MARZO" : va1 = "03"
            Case "ABRIL" : va1 = "04"
            Case "MAYO" : va1 = "05"
            Case "JUNIO" : va1 = "06"
            Case "JULIO" : va1 = "07"
            Case "AGOSTO" : va1 = "08"
            Case "SETIEMBRE" : va1 = "09"
            Case "OCTUBRE" : va1 = "10"
            Case "NOVIEMBRE" : va1 = "11"
            Case "DICIEMBRE" : va1 = "12"
        End Select

        va2 = ComboBox1.SelectedValue.ToString
        'Select Case ComboBox1.Text
        '    Case "GBEDON" : va2 = "0010"
        '    Case "VINCIO" : va2 = "0022"
        '    Case "DBRAVO" : va2 = "0023"
        '    Case "JSALINAS" : va2 = "0025"
        '    Case "GCUEVA" : va2 = "0029"
        '    Case "AMENDO" : va2 = "0026"
        '    Case "VGRAUS" : va2 = "0007"
        '    Case "VPIZARRO" : va2 = "0027"
        '    Case "GBALVIN" : va2 = "0028"
        '    Case "VSILVERIO" : va2 = "0005"
        '    Case "WSALINAS" : va2 = "0034"
        'End Select
        If y = 0 Then
            MsgBox("NO HAY INFORMACION PARA GUARDAR")
        Else
            abrir()
            Dim cmd16 As New SqlCommand("DELETE from comsion_vendedor WHERE ID =@id", conx)
            cmd16.Parameters.AddWithValue("@id", Trim(Label8.Text) + Trim(Label12.Text) + va1 + va2)
            cmd16.ExecuteNonQuery()
            For i = 0 To y - 1
                abrir()
                Dim cmd15 As New SqlCommand(" INSERT INTO comsion_vendedor(id,empresa ,grupo ,cliente ,rubro ,f_pago ,guia,serie,correlativo ,precio_venta ,valor_venta ,venta_descuento ,porcentaje  ,comision ,almacen ,ventas_mes,mes ,vendedor,periodo,fecha,op,apgg) 
                                                                   VALUES (@id,@empresa ,@grupo ,@cliente ,@rubro ,@f_pago ,@guia,@serie,@correlativo ,@precio_venta ,@valor_venta ,@venta_descuento ,@porcentaje  ,@comision ,@almacen ,@ventas_mes,@mes ,@vendedor,@periodo,@fecha,@op,@apgg)", conx)
                cmd15.Parameters.AddWithValue("@id", Trim(Label8.Text) + Trim(Label12.Text) + va1 + va2)
                cmd15.Parameters.AddWithValue("@empresa", Trim(Label12.Text))
                cmd15.Parameters.AddWithValue("@grupo", Trim(DataGridView1.Rows(i).Cells(1).Value))
                cmd15.Parameters.AddWithValue("@cliente", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd15.Parameters.AddWithValue("@rubro", Trim(DataGridView1.Rows(i).Cells(3).Value))
                cmd15.Parameters.AddWithValue("@f_pago", Trim(DataGridView1.Rows(i).Cells(4).Value))
                cmd15.Parameters.AddWithValue("@guia", Trim(DataGridView1.Rows(i).Cells(5).Value))
                cmd15.Parameters.AddWithValue("@serie", Trim(DataGridView1.Rows(i).Cells(6).Value))
                cmd15.Parameters.AddWithValue("@correlativo", Trim(DataGridView1.Rows(i).Cells(7).Value))
                cmd15.Parameters.AddWithValue("@precio_venta", Convert.ToDouble(DataGridView1.Rows(i).Cells(8).Value))
                cmd15.Parameters.AddWithValue("@valor_venta", Convert.ToDouble(DataGridView1.Rows(i).Cells(9).Value))
                cmd15.Parameters.AddWithValue("@venta_descuento", Convert.ToDouble(DataGridView1.Rows(i).Cells(10).Value))
                cmd15.Parameters.AddWithValue("@porcentaje", Convert.ToDouble(DataGridView1.Rows(i).Cells(11).Value))
                cmd15.Parameters.AddWithValue("@comision", Convert.ToDouble(DataGridView1.Rows(i).Cells(12).Value))
                cmd15.Parameters.AddWithValue("@almacen", Trim(DataGridView1.Rows(i).Cells(14).Value))
                cmd15.Parameters.AddWithValue("@ventas_mes", Convert.ToDouble(TextBox5.Text))
                cmd15.Parameters.AddWithValue("@mes", va1)
                cmd15.Parameters.AddWithValue("@vendedor", va2)
                cmd15.Parameters.AddWithValue("@periodo", Trim(Label8.Text))
                If Label15.Text = 1 Then
                    cmd15.Parameters.AddWithValue("@fecha", DataGridView1.Rows(i).Cells(19).Value)
                    cmd15.Parameters.AddWithValue("@op", Trim(DataGridView1.Rows(i).Cells(20).Value))
                Else
                    cmd15.Parameters.AddWithValue("@fecha", DataGridView1.Rows(i).Cells(15).Value)
                    cmd15.Parameters.AddWithValue("@op", Trim(DataGridView1.Rows(i).Cells(16).Value))
                End If
                cmd15.Parameters.AddWithValue("@apgg", "0")
                cmd15.ExecuteNonQuery()
            Next
            MsgBox("LA INFORMACION SE GUARDO CORRECTAMENTE")
            DataGridView1.DataSource = ""
            'TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = "0.00"
            Label5.Text = ""
            TextBox3.Visible = False
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
        End If

    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        Dim valor As Double
        If e.KeyCode = Keys.Enter Then

            valor = Convert.ToDouble(TextBox7.Text) - (5000 * Convert.ToDouble(TextBox6.Text))
            TextBox5.Text = valor.ToString("N2")
        Else
            If e.KeyCode = Keys.Tab Then
                valor = Convert.ToDouble(TextBox7.Text) - (5000 * Convert.ToDouble(TextBox6.Text))
                TextBox5.Text = valor.ToString("N2")
            End If
        End If


    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        NumConFrac2(TextBox6, e)
    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("LA COMISION ESTA SIENDO APROBADA PARA LA IMPRESION DEL REPORTE POR PARTE DEL VENDEDOR  Y SE BLOQUEARIA CUALQUIER MODIFICACION?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Dim va1, va2 As String
            va1 = ""
            va2 = ""

            Select Case ComboBox2.Text
                Case "ENERO" : va1 = "01"
                Case "FEBRERO" : va1 = "02"
                Case "MARZO" : va1 = "03"
                Case "ABRIL" : va1 = "04"
                Case "MAYO" : va1 = "05"
                Case "JUNIO" : va1 = "06"
                Case "JULIO" : va1 = "07"
                Case "AGOSTO" : va1 = "08"
                Case "SETIEMBRE" : va1 = "09"
                Case "OCTUBRE" : va1 = "10"
                Case "NOVIEMBRE" : va1 = "11"
                Case "DICIEMBRE" : va1 = "12"
            End Select

            va2 = ComboBox1.SelectedValue.ToString
            Dim cmd16 As New SqlCommand("update comsion_vendedor set apgg ='1' where id =@id", conx)
            cmd16.Parameters.AddWithValue("@id", Trim(Label8.Text) + Trim(Label12.Text) + va1 + va2)
            cmd16.ExecuteNonQuery()
            MsgBox("SE APROBO LA COMISION")
            DataGridView1.DataSource = ""
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = "0.00"
            Label5.Text = ""
            TextBox3.Visible = False
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
        End If
    End Sub
End Class