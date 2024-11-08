Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Public Class Ingresar_Packing
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Ingresar_Packing_Load(sender As Object, e As EventArgs) Handles MyBase.Load




    End Sub
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim bh As New vpackingtela
        Dim bh5 As New vpackingtela
        Dim ch As New vpackingtela
        Dim bh1 As New fingresopac
        Dim as1, k As Integer
        Dim jh As New Double
        bh.gpartida = TextBox1.Text
        bh.garticulo = TextBox4.Text
        bh.gcodigo_trab = ""
        bh.gnombre_trab = ""
        bh.gnumero_rollos = ComboBox1.Text
        bh.gunidad = "KGS"
        bh.gorden = TextBox12.Text
        If Trim(TextBox13.Text) = "" Then
            bh.gdesidad = 0
        Else
            bh.gdesidad = TextBox13.Text
        End If

        bh.gestado = 1
        k = DataGridView1.Rows.Count
        For o = 0 To k - 1
            jh = jh + DataGridView1.Rows(o).Cells(1).Value
        Next
        bh.garticulo3 = DataGridView1.Rows(0).Cells(4).Value
        bh.gseleccionado = 0
        bh.gtotal = jh
        bh.gROOLO = k
        bh.galmac = Label8.Text

        bh5.gpartida = TextBox1.Text
        bh5.gnumero_rollos = ComboBox1.Text
        bh5.galmac = Label8.Text
        bh1.eliminar_packing(bh5)
        If bh1.insertar_packing_tela(bh) = True Then
            as1 = DataGridView1.Rows.Count
            For i = 0 To as1 - 1
                ch.gnrollo = DataGridView1.Rows(i).Cells(0).Value
                ch.gpeso = DataGridView1.Rows(i).Cells(1).Value
                If Convert.ToString(DataGridView1.Rows(i).Cells(5).Value) = "" Then
                    ch.gposicion = ""

                Else
                    ch.gposicion = DataGridView1.Rows(i).Cells(5).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(3).Value) = "" Then
                    ch.galmacen = ""
                Else
                    ch.galmacen = DataGridView1.Rows(i).Cells(3).Value
                End If



                ch.garticulo2 = DataGridView1.Rows(i).Cells(4).Value
                ch.gpartida = TextBox1.Text
                ch.gestado1 = 1
                ch.gseleccionado = 0
                ch.gidcabe = ComboBox1.Text
                ch.galmac = Label8.Text
                ch.gpeso_neto = DataGridView1.Rows(i).Cells(7).Value
                If Trim(Convert.ToString(DataGridView1.Rows(i).Cells(2).Value)) = "" Then
                    ch.gubic_art = ""
                Else
                    ch.gubic_art = DataGridView1.Rows(i).Cells(2).Value
                End If
                bh1.insertar_packing_tela_detalle(ch)

            Next
            MsgBox("SE REGISTRO CON EXITO")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox9.Text = ""
            TextBox4.Text = ""

            DataGridView1.Rows.Clear()
            DataGridView2.DataSource = ""
            ComboBox1.SelectedIndex = -1
            ComboBox1.Enabled = False
            Button6.Visible = False
            Label16.Visible = False
            TextBox13.Visible = False
            CheckBox1.Checked = False
        Else
            MsgBox("LA PARTIDA YA ESTA REGSTRADA")
        End If


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Dim dt As New DataTable
    Dim jh As New DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim vg As New vpartida
            Dim vg1 As New fpartida



            vg.gpartida = TextBox1.Text

            dt = vg1.buscar_partida(vg)
            If dt.Rows.Count <> 0 Then
                DataGridView2.DataSource = dt
                TextBox3.Text = DataGridView2.Rows(0).Cells(0).Value
                TextBox4.Text = DataGridView2.Rows(0).Cells(1).Value
            Else
                MsgBox("LA PARTIDA INGRESADA NO EXISTE")
            End If

            ComboBox1.Enabled = True
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs)

    End Sub

    'Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        Dim fg As New vpartida
    '        Dim fg1 As New fpartida
    '        Select Case TextBox5.Text.Length
    '            Case "1" : TextBox5.Text = "0000" & "" & TextBox5.Text
    '            Case "2" : TextBox5.Text = "000" & "" & TextBox5.Text
    '            Case "3" : TextBox5.Text = "00" & "" & TextBox5.Text
    '            Case "4" : TextBox5.Text = "0" & "" & TextBox5.Text
    '            Case "5" : TextBox5.Text = TextBox5.Text
    '        End Select
    '        Select Case TextBox5.Text.Length
    '            Case "1" : fg.gcodigo = "0000" & "" & TextBox5.Text
    '            Case "2" : fg.gcodigo = "000" & "" & TextBox5.Text
    '            Case "3" : fg.gcodigo = "00" & "" & TextBox5.Text
    '            Case "4" : fg.gcodigo = "0" & "" & TextBox5.Text
    '            Case "5" : fg.gcodigo = TextBox5.Text
    '        End Select
    '        If fg1.buscar_persona(fg) = "FALSE" Then
    '            MsgBox("EL TARBAJADOR INGRESADO NO EXISTE")
    '        Else
    '            TextBox6.Text = fg1.buscar_persona(fg)
    '            TextBox7.Enabled = True
    '            TextBox7.ReadOnly = False
    '            TextBox9.ReadOnly = False
    '            TextBox10.ReadOnly = False
    '            TextBox2.ReadOnly = False
    '            TextBox9.Select()
    '        End If

    '    End If
    'End Sub




    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            'Using reader As New StreamReader("C:\Program Files (x86)\HMG\WS PESO\peso.txt", Encoding.Default)
            '    TextBox2.Text = reader.ReadToEnd()
            'End Using
            If TextBox2.Text > 60 Or TextBox9.Text = "" Then
                MsgBox("EL PESO INGRESADO EXCEDE EL MAXIMO PERMITIDO O LA ANCHO DE LA TELA ESTA VACIO")
            Else
                Dim k As Integer

                DataGridView1.Rows.Add()
                k = DataGridView1.Rows.Count

                DataGridView1.Rows(k - 1).Cells(0).Value = TextBox7.Text
                DataGridView1.Rows(k - 1).Cells(4).Value = TextBox3.Text
                DataGridView1.Rows(k - 1).Cells(5).Value = TextBox9.Text
                DataGridView1.Rows(k - 1).Cells(3).Value = TextBox10.Text
                DataGridView1.Rows(k - 1).Cells(1).Value = TextBox2.Text
                If CheckBox1.Checked = False Then
                    DataGridView1.Rows(k - 1).Cells(7).Value = TextBox2.Text - 0.3
                Else
                    DataGridView1.Rows(k - 1).Cells(7).Value = TextBox2.Text - 0.4
                End If


                'DataGridView1.CurrentCell = DataGridView1.Item(1, k - 1)

                abrir()
                If CheckBox1.Checked = False Then
                    Dim sql102 As String = "select substring(ncom_4,1,10),substring(ncom_4,11,3) from custom_vianny.dbo.matg040f where partidA_4 ='" + Trim(TextBox1.Text) + "'"
                    Dim cmd102 As New SqlCommand(sql102, conx)
                    Rsr2 = cmd102.ExecuteReader()

                    Rsr2.Read()
                    Rppp_Packig.TextBox1.Text = Rsr2(0)
                    Rppp_Packig.TextBox2.Text = Rsr2(1)
                    Rppp_Packig.TextBox3.Text = TextBox7.Text
                    Rppp_Packig.TextBox4.Text = Convert.ToDouble(TextBox2.Text) - 0.3
                    Rppp_Packig.TextBox5.Text = TextBox10.Text
                    Rppp_Packig.TextBox6.Text = TextBox9.Text
                    Rsr2.Close()

                    Rppp_Packig.Show()
                Else

                    Imp_Expotracion.TextBox1.Text = Trim(TextBox1.Text)
                    Imp_Expotracion.TextBox2.Text = Trim(TextBox7.Text)
                    Imp_Expotracion.TextBox3.Text = Convert.ToDouble(TextBox2.Text) - 0.4
                    Imp_Expotracion.TextBox4.Text = TextBox10.Text
                    Imp_Expotracion.TextBox5.Text = TextBox9.Text
                    Imp_Expotracion.TextBox6.Text = TextBox12.Text
                    Imp_Expotracion.TextBox7.Text = TextBox13.Text
                    Imp_Expotracion.Show()
                End If
                '-------------------------

                Dim I, sum8 As Integer
                Dim cant9, sum9 As Double
                I = DataGridView1.Rows.Count

                For I1 = 0 To I - 1

                    cant9 = Val(DataGridView1.Rows(I1).Cells(7).Value)
                    sum9 = cant9 + Val(sum9)
                    sum8 = sum8 + 1
                Next


                TextBox11.Text = sum8
                TextBox8.Text = sum9
                TextBox7.Text = ""
                TextBox2.Text = ""
                TextBox10.Text = ""
                TextBox7.Select()

            End If


        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        DataGridView1.Rows.Clear()
        DataGridView3.DataSource = ""
        Dim vg2 As New vpackingtela
        Dim vg3 As New fingresopac
        Dim i As Integer

        vg2.gpartida = TextBox1.Text
        vg2.gnumero_rollos = ComboBox1.Text
        If Label7.Text = "LIMA" And ComboBox1.Text = "PRIMERA" Then
            Label8.Text = "10"
            Label18.Text = "1"
        Else


            If Label7.Text = "LIMA" And ComboBox1.Text = "SEGUNDA" Then
                Label8.Text = "10"
                Label18.Text = "2"
            Else
                If Label7.Text = "LIMA" And ComboBox1.Text = "TERCERA" Then
                    Label8.Text = "10"
                    Label18.Text = "4"
                End If
            End If

        End If
        TextBox7.Enabled = True
        TextBox7.ReadOnly = False
        TextBox9.ReadOnly = False
        TextBox10.ReadOnly = False
        TextBox2.ReadOnly = False
        TextBox9.Select()
        vg2.gAL = Trim(Label8.Text)
        jh = vg3.imprimir_packing2(vg2)

        If jh.Rows.Count > 0 Then
            Dim suma10 As Double
            Dim suma2 As Integer

            DataGridView3.DataSource = jh
            Dim UI As Integer
            UI = DataGridView3.Rows.Count
            DataGridView1.Rows.Add(UI)
            For i = 0 To UI - 1
                DataGridView1.Rows(i).Cells(0).Value = DataGridView3.Rows(i).Cells(16).Value
                DataGridView1.Rows(i).Cells(1).Value = DataGridView3.Rows(i).Cells(17).Value
                DataGridView1.Rows(i).Cells(5).Value = DataGridView3.Rows(i).Cells(18).Value
                DataGridView1.Rows(i).Cells(3).Value = DataGridView3.Rows(i).Cells(19).Value
                DataGridView1.Rows(i).Cells(4).Value = DataGridView3.Rows(i).Cells(20).Value
                DataGridView1.Rows(i).Cells(7).Value = DataGridView3.Rows(i).Cells(27).Value
                suma10 = suma10 + DataGridView1.Rows(i).Cells(7).Value
                suma2 = suma2 + 1

            Next

            TextBox8.Text = suma10
            TextBox11.Text = suma2
            TextBox4.Text = DataGridView3.Rows(0).Cells(2).Value
            'TextBox5.Text = DataGridView3.Rows(0).Cells(3).Value
            'TextBox6.Text = DataGridView3.Rows(0).Cells(4).Value
            ComboBox1.Text = DataGridView3.Rows(0).Cells(5).Value
            TextBox12.Text = Trim(DataGridView3.Rows(0).Cells(13).Value)
            TextBox13.Text = DataGridView3.Rows(0).Cells(14).Value
            PictureBox1.Enabled = True

            If TextBox12.Text <> "" Then
                Label16.Visible = True
                TextBox13.Visible = True
                Button6.Visible = True

                CheckBox1.Checked = True


            Else
                Label16.Visible = False
                TextBox13.Visible = False
                Button6.Visible = False
                CheckBox1.Checked = False
            End If
        Else
            'TextBox5.Enabled = True

        End If
        'i = DataGridView1.Rows.Count
        'For a9 = 0 To i - 1
        '    cant9 = Val(DataGridView1.Rows(a9).Cells(1).Value)
        '    sum9 = cant9 + Val(sum9)
        'Next


        'TextBox8.Text = sum9.ToString("N2")
        'TextBox5.Select()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        'Dim asr As New Rpt_Packing
        'asr.TextBox1.Text = TextBox1.Text


    End Sub





    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox7.Text = "" Then
                MsgBox("NO INGRESO INIGUN VALOR EN EL CAMPO ROLLO, VUELVA A INGRESAR")
                TextBox7.Text = ""
                TextBox7.Select()
            Else



                Dim dg As String
                dg = Mid(Year(DateTimePicker1.Value), 3, 2)

                Select Case Trim(TextBox7.Text).Length
                    Case "1" : TextBox7.Text = dg & "0000000" & "" & TextBox7.Text
                    Case "2" : TextBox7.Text = dg & "000000" & "" & TextBox7.Text
                    Case "3" : TextBox7.Text = dg & "00000" & "" & TextBox7.Text
                    Case "4" : TextBox7.Text = dg & "0000" & "" & TextBox7.Text
                    Case "5" : TextBox7.Text = dg & "000" & "" & TextBox7.Text
                    Case "6" : TextBox7.Text = dg & "00" & "" & TextBox7.Text
                    Case "7" : TextBox7.Text = dg & "0" & "" & TextBox7.Text
                    Case "8" : TextBox7.Text = dg & TextBox7.Text

                    Case "10" : TextBox7.Text = TextBox7.Text
                End Select
                Dim y, cont As Integer

                y = DataGridView1.Rows.Count
                cont = 0
                For i = 0 To y - 1
                    If Trim(TextBox7.Text) = Trim(DataGridView1.Rows(i).Cells(0).Value) Then
                        cont = cont + 1
                    End If

                Next
                If cont = "0" Then
                    abrir()
                    '-------------------------
                    Dim sql1022 As String = "select COUNT(rollo_3r) from custom_vianny.dbo.marg0001 where partida_3r ='" + Trim(TextBox1.Text) + "' and rollo_3r ='" + Trim(TextBox7.Text) + "'"
                    Dim cmd1022 As New SqlCommand(sql1022, conx)
                    Rsr22 = cmd1022.ExecuteReader()
                    Rsr22.Read()
                    If Rsr22(0) > 0 Then
                        TextBox2.Select()
                    Else
                        MsgBox("EL ROLLO INGRESADO " + TextBox7.Text + " NO PERTENECE A LA PARTIDA")
                        TextBox7.Text = ""
                        TextBox7.Select()
                    End If
                Else
                    MsgBox("EL ROLLO  " + Trim(TextBox7.Text) + " YA ESTA REGISTRADO, AGREGE OTRO DIFERENTE")
                End If

            End If


        End If
    End Sub

    Dim Rsr2, Rsr22, Rsr3 As SqlDataReader

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, a9, fila As Integer
        Dim cant9, sum9 As Double
        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If CheckBox1.Checked = True Then
            If DataGridView1.Columns(e.ColumnIndex).HeaderText = "KGS BRUTO" Then

                For a9 = 0 To i - 1
                    cant9 = Val(DataGridView1.Rows(a9).Cells(1).Value)
                    sum9 = cant9 + Val(sum9)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(7).Value = Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(1).Value) - 0.4
                Dim u As Integer
                u = e.RowIndex
                TextBox8.Text = sum9.ToString("N2")
            End If
        Else
            If DataGridView1.Columns(e.ColumnIndex).HeaderText = "KGS BRUTO" Then

                For a9 = 0 To i - 1
                    cant9 = Val(DataGridView1.Rows(a9).Cells(1).Value)
                    sum9 = cant9 + Val(sum9)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(7).Value = Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(1).Value) - 0.3
                Dim u As Integer
                u = e.RowIndex
                TextBox8.Text = sum9.ToString("N2")
            End If
        End If




    End Sub

    Private Sub TextBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseDown

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        If TypeOf e.Control Is TextBox Then

            DirectCast(e.Control, TextBox).CharacterCasing = CharacterCasing.Upper

        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        DataGridView1.BeginEdit(True)
    End Sub

    Private Sub TextBox7_MouseDown(sender As Object, e As MouseEventArgs) Handles TextBox7.MouseDown

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i, a9, fila As Integer
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "I" Then

            'For a9 = 0 To i - 1
            '    cant9 = Val(DataGridView1.Rows(a9).Cells(1).Value)
            '    sum9 = cant9 + Val(sum9)
            'Next
            'DataGridView1.Rows(e.RowIndex).Cells(7).Value = Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(1).Value) - 0.3
            If CheckBox1.Checked = False Then
                abrir()
                '-------------------------
                Dim sql102 As String = "select substring(ncom_4,1,10),substring(ncom_4,11,3) from custom_vianny.dbo.matg040f where partidA_4 ='" + Trim(TextBox1.Text) + "'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()

                Rsr2.Read()
                Rppp_Packig.TextBox1.Text = Rsr2(0)
                Rppp_Packig.TextBox2.Text = Rsr2(1)
                Rppp_Packig.TextBox3.Text = Trim(DataGridView1.Rows(fila).Cells(0).Value)
                Rppp_Packig.TextBox4.Text = DataGridView1.Rows(fila).Cells(7).Value
                Rppp_Packig.TextBox5.Text = DataGridView1.Rows(fila).Cells(3).Value
                Rppp_Packig.TextBox6.Text = DataGridView1.Rows(fila).Cells(5).Value
                Rsr2.Close()

                Rppp_Packig.Show()
            Else
                If TextBox13.Text = 0 Or TextBox13.Text = "" Or TextBox12.Text = "" Then
                    MsgBox("ES OBLIGATORIO INGRESAR LA DENSIDAD Y EN NUMERO ORDEN EN EXPORTACION")
                Else
                    Imp_Expotracion.TextBox1.Text = Trim(TextBox1.Text)
                    Imp_Expotracion.TextBox2.Text = DataGridView1.Rows(fila).Cells(0).Value
                    Imp_Expotracion.TextBox3.Text = DataGridView1.Rows(fila).Cells(7).Value
                    Imp_Expotracion.TextBox4.Text = DataGridView1.Rows(fila).Cells(3).Value
                    Imp_Expotracion.TextBox5.Text = DataGridView1.Rows(fila).Cells(5).Value
                    Imp_Expotracion.TextBox6.Text = TextBox12.Text
                    Imp_Expotracion.TextBox7.Text = TextBox13.Text
                    Imp_Expotracion.Show()
                End If

            End If



        End If
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox10.Select()
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox7.Select()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox7.Enabled = True
        TextBox7.ReadOnly = False
        TextBox9.ReadOnly = False
        TextBox10.ReadOnly = False
        TextBox2.ReadOnly = False
        DataGridView1.Enabled = True
        TextBox9.Select()
    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA SALIR DEL PROGRAMA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Me.Close()

        End If
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Label15.Visible = True
            Label16.Visible = True

            TextBox12.Visible = True
            TextBox13.Visible = True
            Button6.Visible = True
            abrir()
            '-------------------------
            Dim sql1023 As String = "SELECT anchor,densidadr FROM CALIDAD_PARTIDA WHERE partida ='" + Trim(TextBox1.Text) + "'"
            Dim cmd1023 As New SqlCommand(sql1023, conx)
            Rsr3 = cmd1023.ExecuteReader()

            If Rsr3.Read() Then
                TextBox9.Text = Rsr3(0)
                TextBox13.Text = Rsr3(1)
                Rsr3.Close()
            Else
                TextBox13.Text = 0
                TextBox9.Text = 0
            End If
            TextBox13.Select()
        Else
            Label15.Visible = False
            Label16.Visible = False
            TextBox12.Visible = False
            TextBox13.Visible = False

            'TextBox9.Visible = False
        End If
    End Sub
    Public Sub NumConFrac(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False

        Else
            e.Handled = True
        End If
    End Sub
    Public Sub NumConFrac1(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
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
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress

        Try
            NumConFrac(TextBox1, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox14.Enabled = True
            TextBox14.Visible = True
            TextBox15.Visible = True
            Label17.Visible = True
            TextBox14.Select()
        Else
            TextBox14.Enabled = False
            TextBox15.Visible = False
            Label17.Visible = False
        End If
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        Try
            NumConFrac(TextBox7, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        Try
            NumConFrac1(TextBox2, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Rpt_export.TextBox1.Text = TextBox1.Text
        Rpt_export.TextBox2.Text = Trim(TextBox12.Text)
        Rpt_export.Show
    End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim k, su As Integer
            Dim dg As String
            dg = Mid(Year(DateTimePicker1.Value), 3, 2)

            Select Case Trim(TextBox14.Text).Length
                Case "1" : TextBox14.Text = dg & "0000000" & "" & TextBox14.Text
                Case "2" : TextBox14.Text = dg & "000000" & "" & TextBox14.Text
                Case "3" : TextBox14.Text = dg & "00000" & "" & TextBox14.Text
                Case "4" : TextBox14.Text = dg & "0000" & "" & TextBox14.Text
                Case "5" : TextBox14.Text = dg & "000" & "" & TextBox14.Text
                Case "6" : TextBox14.Text = dg & "00" & "" & TextBox14.Text
                Case "7" : TextBox14.Text = dg & "0" & "" & TextBox14.Text
                Case "8" : TextBox14.Text = dg & TextBox14.Text

                Case "10" : TextBox14.Text = TextBox14.Text
            End Select
            DataGridView1.Rows.Add()
            k = DataGridView1.Rows.Count
            DataGridView1.Rows(k - 1).Cells(4).Value = TextBox3.Text
            DataGridView1.Rows(k - 1).Cells(0).Value = TextBox14.Text
            DataGridView1.Rows(k - 1).Cells(5).Value = TextBox15.Text
            TextBox14.Text = ""
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        'Using reader As New StreamReader("C:\Program Files (x86)\HMG\WS PESO\peso.wei", Encoding.Default)
        '    TextBox2.Text = reader.ReadToEnd()
        'End Using
        If TextBox2.Text > 50 Or TextBox9.Text = "" Then
            MsgBox("EL PESO INGRESADO EXCEDE EL MAXIMO PERMITIDO O LA ANCHO DE LA TELA ESTA VACIO")
        Else
            Dim k As Integer

            DataGridView1.Rows.Add()
            k = DataGridView1.Rows.Count

            DataGridView1.Rows(k - 1).Cells(0).Value = TextBox7.Text
            DataGridView1.Rows(k - 1).Cells(4).Value = TextBox3.Text
            DataGridView1.Rows(k - 1).Cells(5).Value = TextBox9.Text
            DataGridView1.Rows(k - 1).Cells(3).Value = TextBox10.Text
            DataGridView1.Rows(k - 1).Cells(1).Value = TextBox2.Text
            If CheckBox1.Checked = False Then
                DataGridView1.Rows(k - 1).Cells(7).Value = TextBox2.Text - 0.3
            Else
                DataGridView1.Rows(k - 1).Cells(7).Value = TextBox2.Text - 0.4
            End If


            'DataGridView1.CurrentCell = DataGridView1.Item(1, k - 1)

            abrir()
            If CheckBox1.Checked = False Then
                Dim sql102 As String = "select substring(ncom_4,1,10),substring(ncom_4,11,3) from custom_vianny.dbo.matg040f where partidA_4 ='" + Trim(TextBox1.Text) + "'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()

                Rsr2.Read()
                Rppp_Packig.TextBox1.Text = Rsr2(0)
                Rppp_Packig.TextBox2.Text = Rsr2(1)
                Rppp_Packig.TextBox3.Text = TextBox7.Text
                Rppp_Packig.TextBox4.Text = Convert.ToDouble(TextBox2.Text) - 0.3
                Rppp_Packig.TextBox5.Text = TextBox10.Text
                Rppp_Packig.TextBox6.Text = TextBox9.Text
                Rsr2.Close()

                Rppp_Packig.Show()
            Else

                Imp_Expotracion.TextBox1.Text = Trim(TextBox1.Text)
                Imp_Expotracion.TextBox2.Text = Trim(TextBox7.Text)
                Imp_Expotracion.TextBox3.Text = Convert.ToDouble(TextBox2.Text) - 0.4
                Imp_Expotracion.TextBox4.Text = TextBox10.Text
                Imp_Expotracion.TextBox5.Text = TextBox9.Text
                Imp_Expotracion.TextBox6.Text = TextBox12.Text
                Imp_Expotracion.TextBox7.Text = TextBox13.Text
                Imp_Expotracion.Show()
            End If
            '-------------------------

            Dim I, sum8 As Integer
            Dim cant9, sum9 As Double
            I = DataGridView1.Rows.Count

            For I1 = 0 To I - 1

                cant9 = Val(DataGridView1.Rows(I1).Cells(7).Value)
                sum9 = cant9 + Val(sum9)
                sum8 = sum8 + 1
            Next


            TextBox11.Text = sum8
            TextBox8.Text = sum9
            TextBox7.Text = ""
            TextBox2.Text = ""
            TextBox10.Text = ""
            TextBox7.Select()

        End If

    End Sub
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

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        abrir()
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("SE ELIMINAR TODA LA INFORMACIO DEL PACKING?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Dim agregar As String = " DELETE FROM cab_ingresop WHERE partida ='" + Trim(TextBox1.Text) + "' and numero_rollos ='" + Trim(ComboBox1.Text) + "' AND almac = '" + Label8.Text + "'"
            Dim agregar1 As String = " DELETE FROM det_ingresop where partida='" + Trim(TextBox1.Text) + "' and id_cabe ='" + Trim(ComboBox1.Text) + "' AND almac = '" + Label8.Text + "'"
            ELIMINAR(agregar)
            ELIMINAR(agregar1)
            MsgBox("SE ELIMINO LA INFORMACION CORRECTAMENTE")
            Me.Close()
        End If


    End Sub

    Private Sub TextBox2_PaddingChanged(sender As Object, e As EventArgs) Handles TextBox2.PaddingChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub

    Private Sub TextBox7_ReadOnlyChanged(sender As Object, e As EventArgs) Handles TextBox7.ReadOnlyChanged

    End Sub

    Private Sub TextBox2_MouseLeave(sender As Object, e As EventArgs) Handles TextBox2.MouseLeave

    End Sub
End Class