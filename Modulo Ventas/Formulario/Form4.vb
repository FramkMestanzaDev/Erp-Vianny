Imports System.Data.SqlClient
Public Class Form4
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
    Dim DT, dt2, dt3 As New DataTable
    Dim Rs11 As SqlDataReader
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then


            DataGridView5.DataSource = ""

                DataGridView1.Rows.Clear()
                dt2.Clear()
                DT.Clear()
                Dim fg As New fpartida
                Dim fg1 As New vpartida
                Dim fg2 As New vpartida
                Dim j As Integer
                Dim V, A, R, p As Integer
                V = 0
                A = 0
                R = 0
                p = 0
                fg1.gpartida = TextBox1.Text

                dt2 = fg.buscar_partida_ingresada(fg1)
                DataGridView2.DataSource = dt2
                j = DataGridView2.Rows.Count
                If j > 0 Then
                    Label21.Text = 1
                    TextBox2.Text = DataGridView2.Rows(0).Cells(1).Value
                    TextBox4.Text = DataGridView2.Rows(0).Cells(2).Value
                    TextBox13.Text = DataGridView2.Rows(0).Cells(3).Value
                    TextBox3.Text = DataGridView2.Rows(0).Cells(4).Value
                    fg2.gpartida = TextBox1.Text
                    TextBox10.Text = fg.buscar_rollos(fg2)
                    TextBox16.Enabled = True
                TextBox5.Text = Trim(DataGridView2.Rows(0).Cells(6).Value)
                TextBox6.Text = Trim(DataGridView2.Rows(0).Cells(7).Value)
                TextBox7.Text = Trim(DataGridView2.Rows(0).Cells(8).Value)
                TextBox8.Text = Trim(DataGridView2.Rows(0).Cells(9).Value)
                TextBox9.Text = Trim(DataGridView2.Rows(0).Cells(10).Value)
                TextBox11.Text = Trim(DataGridView2.Rows(0).Cells(11).Value)
                    TextBox12.Text = Trim(DataGridView2.Rows(0).Cells(12).Value)
                    TextBox14.Text = DataGridView2.Rows(0).Cells(13).Value
                TextBox15.Text = Trim(DataGridView2.Rows(0).Cells(16).Value)
                TextBox24.Text = DataGridView2.Rows(0).Cells(18).Value
                    TextBox25.Text = DataGridView2.Rows(0).Cells(19).Value
                TextBox26.Text = Trim(DataGridView2.Rows(0).Cells(20).Value)
                ComboBox1.Text = Trim(DataGridView2.Rows(0).Cells(21).Value)
                ComboBox2.Text = Trim(DataGridView2.Rows(0).Cells(22).Value)
                ComboBox3.Text = Trim(DataGridView2.Rows(0).Cells(23).Value)
                If DataGridView2.Rows(0).Cells(14).Value = 1 Then
                        CheckBox1.Checked = False
                    Else
                        CheckBox1.Checked = True
                    End If

            If Convert.ToString(DataGridView2.Rows(0).Cells(24).Value) <> "" Then
                DataGridView1.Rows.Add(j)
                For i = 0 To j - 1
                    DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(25).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(i).Cells(26).Value
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView2.Rows(i).Cells(27).Value
                    DataGridView1.Rows(i).Cells(3).Value = DataGridView2.Rows(i).Cells(28).Value
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(29).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(30).Value
                    DataGridView1.Rows(i).Cells(6).Value = DataGridView2.Rows(i).Cells(31).Value
                    DataGridView1.Rows(i).Cells(7).Value = DataGridView2.Rows(i).Cells(32).Value
                    DataGridView1.Rows(i).Cells(8).Value = DataGridView2.Rows(i).Cells(33).Value
                    DataGridView1.Rows(i).Cells(9).Value = DataGridView2.Rows(i).Cells(34).Value
                    DataGridView1.Rows(i).Cells(10).Value = DataGridView2.Rows(i).Cells(35).Value
                    DataGridView1.Rows(i).Cells(11).Value = DataGridView2.Rows(i).Cells(36).Value
                    DataGridView1.Rows(i).Cells(12).Value = DataGridView2.Rows(i).Cells(37).Value
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView2.Rows(i).Cells(64).Value
                    DataGridView1.Rows(i).Cells(14).Value = DataGridView2.Rows(i).Cells(38).Value
                    DataGridView1.Rows(i).Cells(15).Value = DataGridView2.Rows(i).Cells(39).Value
                    DataGridView1.Rows(i).Cells(16).Value = DataGridView2.Rows(i).Cells(40).Value
                    DataGridView1.Rows(i).Cells(17).Value = DataGridView2.Rows(i).Cells(41).Value
                    DataGridView1.Rows(i).Cells(18).Value = DataGridView2.Rows(i).Cells(42).Value
                    DataGridView1.Rows(i).Cells(19).Value = DataGridView2.Rows(i).Cells(43).Value
                    DataGridView1.Rows(i).Cells(20).Value = DataGridView2.Rows(i).Cells(44).Value
                    DataGridView1.Rows(i).Cells(21).Value = DataGridView2.Rows(i).Cells(45).Value
                    DataGridView1.Rows(i).Cells(22).Value = DataGridView2.Rows(i).Cells(46).Value
                    DataGridView1.Rows(i).Cells(23).Value = DataGridView2.Rows(i).Cells(47).Value
                    DataGridView1.Rows(i).Cells(24).Value = DataGridView2.Rows(i).Cells(48).Value
                    DataGridView1.Rows(i).Cells(25).Value = DataGridView2.Rows(i).Cells(49).Value
                    DataGridView1.Rows(i).Cells(26).Value = DataGridView2.Rows(i).Cells(50).Value
                    DataGridView1.Rows(i).Cells(27).Value = DataGridView2.Rows(i).Cells(51).Value
                    DataGridView1.Rows(i).Cells(28).Value = DataGridView2.Rows(i).Cells(52).Value
                    DataGridView1.Rows(i).Cells(29).Value = DataGridView2.Rows(i).Cells(53).Value
                    DataGridView1.Rows(i).Cells(30).Value = DataGridView2.Rows(i).Cells(54).Value
                    DataGridView1.Rows(i).Cells(31).Value = DataGridView2.Rows(i).Cells(55).Value
                    DataGridView1.Rows(i).Cells(32).Value = DataGridView2.Rows(i).Cells(65).Value
                    DataGridView1.Rows(i).Cells(33).Value = DataGridView2.Rows(i).Cells(56).Value
                    DataGridView1.Rows(i).Cells(34).Value = DataGridView2.Rows(i).Cells(57).Value
                    DataGridView1.Rows(i).Cells(35).Value = DataGridView2.Rows(i).Cells(58).Value
                    DataGridView1.Rows(i).Cells(36).Value = DataGridView2.Rows(i).Cells(59).Value
                    DataGridView1.Rows(i).Cells(37).Value = DataGridView2.Rows(i).Cells(60).Value
                    DataGridView1.Rows(i).Cells(38).Value = DataGridView2.Rows(i).Cells(61).Value
                    Dim uy As Integer
                        If Trim(DataGridView2.Rows(i).Cells(56).Value) = "" Then
                            uy = 0
                        Else
                            uy = Trim(DataGridView2.Rows(i).Cells(56).Value)
                        End If
                    If uy = 1 Then

                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Green
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                        V = V + 1
                    Else
                        If uy = 2 Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                            DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.DarkRed
                            A = A + 1
                        Else
                            If uy = 3 Then
                                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                R = R + 1
                            Else
                                If uy = 0 Then
                                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                                    p = p + 1
                                End If
                            End If
                        End If


                    End If
                Next
                TextBox27.Text = V
                TextBox28.Text = A
                TextBox29.Text = R
                TextBox30.Text = p
            End If
        Else
                    DT = fg.buscar_partida_calidad(fg1)
                    If DT.Rows.Count <> 0 Then
                        DataGridView5.DataSource = DT
                        If Convert.ToString(DataGridView5.Rows(0).Cells(3).Value) = "" Then
                            TextBox2.Text = ""
                        Else
                            TextBox2.Text = DataGridView5.Rows(0).Cells(3).Value
                        End If
                        If Convert.ToString(DataGridView5.Rows(0).Cells(4).Value) = "" Then
                            TextBox3.Text = ""
                        Else
                            TextBox3.Text = DataGridView5.Rows(0).Cells(4).Value
                        End If
                        If Convert.ToString(DataGridView5.Rows(0).Cells(2).Value) = "" Then
                            TextBox4.Text = ""
                        Else
                            TextBox4.Text = DataGridView5.Rows(0).Cells(2).Value
                        End If
                        If Convert.ToString(DataGridView5.Rows(0).Cells(1).Value) = "" Then
                            TextBox13.Text = ""
                        Else
                            TextBox13.Text = DataGridView5.Rows(0).Cells(1).Value
                        End If

                        fg2.gpartida = TextBox1.Text
                        TextBox10.Text = fg.buscar_rollos(fg2)
                        TextBox16.Enabled = True
                        'DataGridView3.DataSource = dt3
                        Label21.Text = 0
                        TextBox27.Text = 0
                        TextBox28.Text = 0
                        TextBox29.Text = 0
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox15.Text = ""
                        TextBox7.Text = ""
                        TextBox8.Text = ""
                        TextBox9.Text = ""
                        TextBox11.Text = ""
                        TextBox12.Text = ""
                        TextBox24.Text = 0
                        abrir()
                        Dim sql11 As String = "SELECT DISTINCT maquina_3r,lote_3r FROM custom_vianny.dbo.marg0001 WHERE partida_3r ='" + Trim(TextBox1.Text) + "'"
                        Dim cmd11 As New SqlCommand(sql11, conx)
                        Rs11 = cmd11.ExecuteReader()
                        Rs11.Read()
                        TextBox25.Text = Rs11(0)
                        TextBox26.Text = Rs11(1)
                        Rs11.Close()

                        Dim cmd As New SqlDataAdapter("SELECT rollo_3r FROM   custom_vianny.dbo.Marg0001 WHERE partida_3r='" + Trim(TextBox1.Text) + "' AND flag_3r >0", conx)

                        Dim da13 As New DataTable

                        cmd.Fill(da13)
                        If da13.Rows.Count > 0 Then
                            DataGridView6.DataSource = da13
                            Dim OP As New Integer
                            OP = DataGridView6.Rows.Count
                            DataGridView1.Rows.Add(OP)
                            For ui = 0 To OP - 1

                                DataGridView1.Rows(ui).Cells(38).Value = Trim(DataGridView6.Rows(ui).Cells(0).Value)

                            Next
                        End If
                    Else
                        MsgBox("LA PARTIDA NO EXISTE")
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox13.Text = ""
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox7.Text = ""
                        TextBox8.Text = ""
                        TextBox9.Text = ""
                        TextBox11.Text = ""
                        TextBox12.Text = ""
                        TextBox14.Text = ""
                        TextBox1.Text = ""
                        CheckBox1.Checked = False
                    End If
                End If
            If Label22.Text = 2 Then
                DataGridView1.ReadOnly = True
                TextBox5.Enabled = False
                TextBox6.Enabled = False
                TextBox7.Enabled = False
                TextBox8.Enabled = False
                TextBox9.Enabled = False
                TextBox15.Enabled = False
                TextBox11.Enabled = False
                TextBox12.Enabled = False
                TextBox24.Enabled = False
                Button3.Enabled = False
                Button2.Enabled = False
                Button1.Enabled = False
                TextBox1.ReadOnly = True
                Button4.Enabled = False
            End If



        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim archivo As New OpenFileDialog
            archivo.Filter = "Archivo excel|*.xlsx"
            If archivo.ShowDialog = DialogResult.OK Then
                TextBox14.Text = archivo.FileName

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        DataGridView1.ColumnHeadersHeight = 150
        TextBox17.Text = Mid(Year(DateTimePicker1.Value), 3, 2)
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Dim ad As New Calidad
        ad.TextBox1.Text = TextBox1.Text
        ad.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Det_Rollo.Label1.Text = "AUDITOR"
        Det_Rollo.Label2.Text = 31
        Det_Rollo.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'If e.ColumnIndex = 37 Then
        '    MsgBox("bien")
        '    ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
        '    Dim num1, num2 As Integer
        '    num1 = e.RowIndex.ToString
        '    num2 = e.ColumnIndex.ToString
        '    Dim y1 As Integer
        '    y1 = DataGridView1.Rows.Count
        '    For ol = 0 To y1 - 1
        '        If DataGridView1.Rows(ol).Cells(37).Value = True Then
        '            DataGridView1.Rows(ol).DefaultCellStyle.BackColor = Color.Coral
        '        Else
        '            DataGridView1.Rows(ol).DefaultCellStyle.BackColor = Color.White
        '        End If
        '    Next

        'End If
    End Sub

    Private Sub DataGridView1_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
        If e.RowIndex = -1 Then
            Dim i As Integer = 0
            For i = 0 To 35
                If e.ColumnIndex = i Then

                    e.PaintBackground(e.CellBounds, True)
                    e.Graphics.TranslateTransform(e.CellBounds.Left, e.CellBounds.Bottom)
                    e.Graphics.RotateTransform(270)
                    e.Graphics.DrawString(e.FormattedValue, e.CellStyle.Font, Brushes.Black, 5, 5)
                    e.Graphics.ResetTransform()

                    e.Handled = True
                End If
            Next
        End If
    End Sub

    Private Sub TextBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim fg2 As New vpartida
            Dim fg As New fpartida
            Dim sum As Integer
            sum = 0

            Select Case TextBox16.Text.Length
                Case "1" : TextBox16.Text = "0000000" & "" & TextBox16.Text
                Case "2" : TextBox16.Text = "000000" & "" & TextBox16.Text
                Case "3" : TextBox16.Text = "00000" & "" & TextBox16.Text
                Case "4" : TextBox16.Text = "0000" & "" & TextBox16.Text
                Case "5" : TextBox16.Text = "000" & "" & TextBox16.Text
                Case "6" : TextBox16.Text = "00" & "" & TextBox16.Text
                Case "7" : TextBox16.Text = "0" & "" & TextBox16.Text
                Case "8" : TextBox16.Text = TextBox16.Text
            End Select
            fg2.gpartida = TextBox1.Text
            fg2.grollo = Trim(TextBox17.Text) & "" & Trim(TextBox16.Text)

            sum = fg.buscar_rollos_existe(fg2)
            If sum > 0 Then
                DataGridView1.Rows.Add(1)
                Dim ui As Integer
                ui = DataGridView1.Rows.Count
                If ui = 1 Then
                    DataGridView1.Rows(0).Cells(36).Value = TextBox17.Text & "" & TextBox16.Text
                Else
                    DataGridView1.Rows(ui - 1).Cells(36).Value = TextBox17.Text & "" & TextBox16.Text
                End If
                TextBox16.Text = ""
            Else
                MsgBox("EL ROLLO SOLICITADO NO EXISTE")
            End If

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim gj As New vcalidadpartida
        Dim gj1 As New fpartida
        Dim gj2 As New vpartida
        Dim jj As New vdetallecalidad
        Dim l As Integer
        gj.gpartida = TextBox1.Text
        gj.gcodigo = TextBox2.Text
        gj.gcolor_des = TextBox4.Text
        gj.gcolor_cod = TextBox13.Text
        gj.gdescripcion = TextBox3.Text
        gj.grollos = TextBox10.Text
        gj.ganchor = TextBox5.Text
        gj.gdensidadr = TextBox6.Text
        gj.glavadoa = TextBox7.Text
        gj.glavadol = TextBox8.Text
        gj.glavador = TextBox9.Text
        gj.gobservacion = TextBox11.Text
        gj.grevisado = TextBox12.Text
        gj.gadjuntado = TextBox14.Text
        gj.gelongacion = TextBox15.Text
        gj.gmaquina = TextBox25.Text
        gj.glote = TextBox26.Text
        gj.gCENTRO_ORILLO = ComboBox1.Text
        gj.gBARRADURA = ComboBox3.Text
        gj.gDEGRADE = ComboBox2.Text
        If CheckBox1.Checked = True Then
            gj.greproceso = 2
        Else
            gj.greproceso = 1
        End If
        gj.gfecha = DateTimePicker1.Value
        gj.gestado = 1
        gj2.gpartida = Trim(TextBox1.Text)
        gj.gmerma = TextBox24.Text
        gj1.eliminar_calidad_partida(gj2)
        gj1.insertar_calidad_partida(gj)
        l = DataGridView1.Rows.Count
        For i = 0 To l - 1
            If Convert.ToString(DataGridView1.Rows(i).Cells(0).Value) = "" Then
                jj.gDHCFM = ""
            Else
                jj.gDHCFM = DataGridView1.Rows(i).Cells(0).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(1).Value) = "" Then
                jj.gDHCPp = ""
            Else
                jj.gDHCPp = DataGridView1.Rows(i).Cells(1).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(2).Value) = "" Then
                jj.gDHCHG = ""
            Else
                jj.gDHCHG = DataGridView1.Rows(i).Cells(2).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(3).Value) = "" Then
                jj.gDHBF = ""
            Else
                jj.gDHBF = DataGridView1.Rows(i).Cells(3).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                jj.gDHTIH = ""
            Else
                jj.gDHTIH = DataGridView1.Rows(i).Cells(4).Value
            End If

            If Convert.ToString(DataGridView1.Rows(i).Cells(5).Value) = "" Then
                jj.gDTR = ""
            Else
                jj.gDTR = DataGridView1.Rows(i).Cells(5).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(6).Value) = "" Then
                jj.gDTLC = ""
            Else
                jj.gDTLC = DataGridView1.Rows(i).Cells(6).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(7).Value) = "" Then
                jj.gDTEL = ""
            Else
                jj.gDTEL = DataGridView1.Rows(i).Cells(7).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(8).Value) = "" Then
                jj.gDTRL = ""
            Else
                jj.gDTRL = DataGridView1.Rows(i).Cells(8).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(9).Value) = "" Then
                jj.gDTGA = ""
            Else
                jj.gDTGA = DataGridView1.Rows(i).Cells(9).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(10).Value) = "" Then
                jj.gDTLA = ""
            Else
                jj.gDTLA = DataGridView1.Rows(i).Cells(10).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(11).Value) = "" Then
                jj.gDTAG = ""
            Else
                jj.gDTAG = DataGridView1.Rows(i).Cells(11).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(12).Value) = "" Then
                jj.gDTCA = ""
            Else
                jj.gDTCA = DataGridView1.Rows(i).Cells(12).Value
            End If
            '***************************************
            If Convert.ToString(DataGridView1.Rows(i).Cells(13).Value) = "" Then
                jj.gmotas = ""
            Else
                jj.gmotas = DataGridView1.Rows(i).Cells(13).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(14).Value) = "" Then
                jj.gDTCT = ""
            Else
                jj.gDTCT = DataGridView1.Rows(i).Cells(14).Value
            End If
            '********************************
            If Convert.ToString(DataGridView1.Rows(i).Cells(15).Value) = "" Then
                jj.gDTAP = ""
            Else
                jj.gDTAP = DataGridView1.Rows(i).Cells(15).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(16).Value) = "" Then
                jj.gDTALR = ""
            Else
                jj.gDTALR = DataGridView1.Rows(i).Cells(16).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(17).Value) = "" Then
                jj.gDTATM = ""
            Else
                jj.gDTATM = DataGridView1.Rows(i).Cells(17).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(18).Value) = "" Then
                jj.gDTAML = ""
            Else
                jj.gDTAML = DataGridView1.Rows(i).Cells(18).Value
            End If


            If Convert.ToString(DataGridView1.Rows(i).Cells(19).Value) = "" Then
                jj.gDTAMV = ""
            Else
                jj.gDTAMV = DataGridView1.Rows(i).Cells(19).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(20).Value) = "" Then
                jj.gDTANMI = ""
            Else
                jj.gDTANMI = DataGridView1.Rows(i).Cells(20).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(21).Value) = "" Then
                jj.gDTAQT = ""
            Else
                jj.gDTAQT = DataGridView1.Rows(i).Cells(21).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(22).Value) = "" Then
                jj.gDTAPMS = ""
            Else
                jj.gDTAPMS = DataGridView1.Rows(i).Cells(22).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(23).Value) = "" Then
                jj.gDTARMB = ""
            Else
                jj.gDTARMB = DataGridView1.Rows(i).Cells(23).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(24).Value) = "" Then
                jj.gDTAMC = ""
            Else
                jj.gDTAMC = DataGridView1.Rows(i).Cells(24).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(25).Value) = "" Then
                jj.gDTAMG = ""
            Else
                jj.gDTAMG = DataGridView1.Rows(i).Cells(25).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(26).Value) = "" Then
                jj.gDTAMS = ""
            Else
                jj.gDTAMS = DataGridView1.Rows(i).Cells(26).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(27).Value) = "" Then
                jj.gDTAMCA = ""
            Else
                jj.gDTAMCA = DataGridView1.Rows(i).Cells(27).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(28).Value) = "" Then
                jj.gDTAM = ""
            Else
                jj.gDTAM = DataGridView1.Rows(i).Cells(28).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(29).Value) = "" Then
                jj.gDTAAG = ""
            Else
                jj.gDTAAG = DataGridView1.Rows(i).Cells(29).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(30).Value) = "" Then
                jj.gDTALMT = ""
            Else
                jj.gDTALMT = DataGridView1.Rows(i).Cells(30).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(31).Value) = "" Then
                jj.gDTALQ = ""

            Else
                jj.gDTALQ = DataGridView1.Rows(i).Cells(31).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(32).Value) = "" Then
                jj.gcrama = ""
            Else
                jj.gcrama = DataGridView1.Rows(i).Cells(32).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(33).Value) = "" Then
                jj.gCC = ""
            Else
                jj.gCC = DataGridView1.Rows(i).Cells(33).Value
            End If

            If Convert.ToString(DataGridView1.Rows(i).Cells(34).Value) = "" Then
                jj.gCE = ""
            Else
                jj.gCE = DataGridView1.Rows(i).Cells(34).Value
            End If


            If Convert.ToString(DataGridView1.Rows(i).Cells(35).Value) = "" Then
                jj.gCDT = ""
            Else
                jj.gCDT = DataGridView1.Rows(i).Cells(35).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(36).Value) = "" Then
                jj.gDAR = ""
            Else
                jj.gDAR = DataGridView1.Rows(i).Cells(36).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(37).Value) = "" Then
                jj.gDo1 = ""
            Else
                jj.gDo1 = DataGridView1.Rows(i).Cells(37).Value
            End If
            If Convert.ToString(DataGridView1.Rows(i).Cells(38).Value) = "" Then
                jj.gROLLO = ""
            Else
                jj.gROLLO = DataGridView1.Rows(i).Cells(38).Value
            End If
            Dim uy As Integer
            If Trim(DataGridView1.Rows(i).Cells(33).Value) = "" Then
                uy = 0
            Else
                uy = Trim(DataGridView1.Rows(i).Cells(33).Value)
            End If
            If uy = 1 Then
                jj.gestado = 1
            Else
                If uy = 2 Then
                    jj.gestado = 2
                Else
                    If uy = 3 Then
                        jj.gestado = 3
                    Else
                        If uy = 0 Then
                            jj.gestado = ""
                        End If
                    End If

                End If


            End If

            jj.gPARTIDA = TextBox1.Text
            gj1.insertar_DETALLE_calidad_partida(jj)
        Next



        MsgBox("LA INFORMCION SE REGISRO CORRECTAMENTE")

        Dim respuesta12 As DialogResult
        respuesta12 = MessageBox.Show("DESEA CONTINUAR INGRESANDO INFORMACION?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta12 = Windows.Forms.DialogResult.Yes) Then

        Else
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox13.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox11.Text = ""
            TextBox12.Text = ""
            TextBox14.Text = ""
            TextBox1.Text = ""
            CheckBox1.Checked = False
            DataGridView1.Rows.Clear()
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim respuesta As DialogResult
        Dim TAB As Integer
        TAB = DataGridView1.Rows.Count
        If TAB <> 0 Then
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)


            End If
        Else
            MsgBox("NO HAY PRODUCTOS A ELIMINAR")
        End If
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

    End Sub
    Dim da As New DataTable
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        abrir()
        Dim cmd As New SqlDataAdapter("select * from CALIDAD_PARTIDA c inner join DETALLE_CALIDAD_PARTIDA d on c.partida = d.PARTIDA where c.partida ='" + Trim(TextBox1.Text) + "'", conx)

        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView2.DataSource = da
        Else
            DataGridView2.DataSource = Nothing
        End If

        Dim ex As New Exportar
        ex.llenarExcel(DataGridView2)
    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit

    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        DataGridView1.BeginEdit(True)
        Dim uy As Integer

        If Label21.Text = 1 Then
            Dim J As Integer

            J = DataGridView1.Rows.Count
            For I = 0 To J - 1
                'MsgBox(Trim(DataGridView1.Rows(I).Cells(33).Value))
                If Me.DataGridView1.CurrentRow.Index.ToString() = I Then

                    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Cyan
                    DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.Black
                Else
                    If Trim(DataGridView2.Rows(I).Cells(53).Value) = "" Then
                        uy = 0
                    Else
                        uy = Trim(DataGridView2.Rows(I).Cells(53).Value)
                    End If
                    If uy = 1 Then

                        DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Green
                        DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White
                    Else
                        If uy = 2 Then
                            DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Yellow
                            DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.DarkRed
                        Else
                            If uy = 3 Then
                                DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Red
                                DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White
                            Else
                                If uy = 0 Then
                                    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.White
                                    DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.Black
                                End If



                            End If
                        End If


                    End If
                End If
            Next
        Else
            Dim J As Integer

            J = DataGridView1.Rows.Count
            For I = 0 To J - 1
                'MsgBox(Trim(DataGridView1.Rows(I).Cells(33).Value))
                If Me.DataGridView1.CurrentRow.Index.ToString() = I Then

                    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Cyan
                Else

                    'If Trim(DataGridView1.Rows(I).Cells(33).Value) = 1 Then

                    '    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Green
                    '    DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White
                    'Else
                    '    If Trim(DataGridView1.Rows(I).Cells(33).Value) = 2 Then
                    '        DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Yellow
                    '        DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.DarkRed
                    '    Else
                    '        If Trim(DataGridView1.Rows(I).Cells(33).Value) = 3 Then
                    '            DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Red
                    '            DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White
                    '        Else
                    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.White
                    DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.Black
                    'End If
                    'End If


                    'End If
                End If
            Next
        End If




    End Sub
End Class