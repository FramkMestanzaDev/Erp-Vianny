Imports System.Data.SqlClient
Imports System.Net.Mail

Public Class pesaod_rollo
    Dim dt As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Dim ty, ty1 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then

            If TextBox1.Text = "" Then
                MsgBox("NO INGRESO NINGUN ROLLO")
                TextBox1.Select()
            Else
                Dim respuesta As DialogResult
                Dim l As New vtejeduria
                Dim kl As New ftejeduria
                Dim p As New vpesadorollo
                Dim dg As String
                dg = Mid(Year(DateTimePicker2.Value), 3, 2)
                Select Case TextBox1.Text.Length
                    Case "1" : TextBox1.Text = dg & "0000000" & "" & TextBox1.Text
                    Case "2" : TextBox1.Text = dg & "000000" & "" & TextBox1.Text
                    Case "3" : TextBox1.Text = dg & "00000" & "" & TextBox1.Text
                    Case "4" : TextBox1.Text = dg & "0000" & "" & TextBox1.Text
                    Case "5" : TextBox1.Text = dg & "000" & "" & TextBox1.Text
                    Case "6" : TextBox1.Text = dg & "00" & "" & TextBox1.Text
                    Case "7" : TextBox1.Text = dg & "0" & "" & TextBox1.Text
                    Case "8" : TextBox1.Text = dg & TextBox1.Text

                    Case "10" : TextBox1.Text = TextBox1.Text
                End Select
                l.gccia = Label28.Text
                l.grolloini = TextBox1.Text
                l.grollofin = TextBox1.Text

                dt = kl.reporte_pesado_rollo(l)

                If dt.Rows.Count <> 0 Then
                    DataGridView1.DataSource = dt
                    TextBox7.Text = DataGridView1.Rows(0).Cells(8).Value
                    TextBox8.Text = DataGridView1.Rows(0).Cells(44).Value
                    TextBox9.Text = Replace(DataGridView1.Rows(0).Cells(7).Value, "_", " ")
                    TextBox10.Text = DataGridView1.Rows(0).Cells(28).Value
                    TextBox11.Text = DataGridView1.Rows(0).Cells(25).Value
                    TextBox13.Text = DataGridView1.Rows(0).Cells(5).Value
                    TextBox12.Text = DataGridView1.Rows(0).Cells(10).Value
                    TextBox14.Text = DataGridView1.Rows(0).Cells(14).Value
                    TextBox16.Text = DataGridView1.Rows(0).Cells(9).Value
                    TextBox15.Text = DataGridView1.Rows(0).Cells(45).Value
                    TextBox19.Text = DataGridView1.Rows(0).Cells(6).Value
                    TextBox22.Text = DataGridView1.Rows(0).Cells(21).Value
                    TextBox23.Text = DataGridView1.Rows(0).Cells(22).Value
                    TextBox20.Text = DataGridView1.Rows(0).Cells(42).Value
                    TextBox21.Text = DataGridView1.Rows(0).Cells(40).Value
                    TextBox17.Text = DataGridView1.Rows(0).Cells(46).Value
                    TextBox18.Text = DataGridView1.Rows(0).Cells(39).Value
                    TextBox24.Text = DataGridView1.Rows(0).Cells(11).Value
                    TextBox25.Text = DataGridView1.Rows(0).Cells(1).Value
                    DateTimePicker3.Value = DataGridView1.Rows(0).Cells(29).Value
                    TextBox6.Text = Trim(DataGridView1.Rows(0).Cells(20).Value)
                    TextBox2.Text = DataGridView1.Rows(0).Cells(51).Value
                    TextBox3.Text = Trim(DataGridView1.Rows(0).Cells(30).Value)
                    If Trim(TextBox3.Text).Length = 0 Then
                        TextBox4.Text = ""
                    Else
                        p.gtejedor = TextBox3.Text
                        TextBox4.Text = Trim(kl.MOSTRAR_TEJEDOR(p))
                    End If
                    Select Case Trim(DataGridView1.Rows(0).Cells(19).Value)
                        Case "18" : ComboBox2.SelectedIndex = 0
                        Case "20" : ComboBox2.SelectedIndex = 1
                        Case "24" : ComboBox2.SelectedIndex = 2
                        Case "28" : ComboBox2.SelectedIndex = 3
                        Case "32" : ComboBox2.SelectedIndex = 4
                        Case "40" : ComboBox2.SelectedIndex = 5


                    End Select
                    Select Case Trim(DataGridView1.Rows(0).Cells(24).Value)
                        Case "90" : ComboBox3.SelectedIndex = 0
                        Case "96" : ComboBox3.SelectedIndex = 1


                    End Select
                    TextBox5.Text = DataGridView1.Rows(0).Cells(16).Value
                    TextBox3.Select()

                    If Trim(DataGridView1.Rows(0).Cells(15).Value) = "1" Then
                        ComboBox1.SelectedIndex = 0
                    Else
                        If Trim(DataGridView1.Rows(0).Cells(15).Value) = "2" Then
                            ComboBox1.SelectedIndex = 1
                        Else
                            ComboBox1.SelectedIndex = 2
                        End If

                    End If
                Else
                    If Hour(DateTimePicker4.Value) >= "07" And Hour(DateTimePicker4.Value) <= "19" Then
                        ComboBox1.SelectedIndex = 0
                    Else
                        ComboBox1.SelectedIndex = 1
                    End If
                End If
                TextBox3.Select()
                Dim KP As Double
                KP = TextBox6.Text
                TextBox6.Text = Format(KP, "##,##00.00")

                TextBox6.Enabled = True

                If Trim(DataGridView1.Rows(0).Cells(18).Value) = 0 Then
                    TextBox6.Enabled = False
                    TextBox3.Enabled = False
                    Button1.Enabled = False
                    Label33.Visible = True
                    ComboBox1.Enabled = False
                    ComboBox2.Enabled = False
                    ComboBox3.Enabled = False
                    TextBox3.Select()
                Else
                    If Trim(DataGridView1.Rows(0).Cells(46).Value) = "AUDITADO" Or Trim(DataGridView1.Rows(0).Cells(46).Value) = "ACTIVO" Then
                        'MsgBox("EL ROLLO ESTA AUDITADO NO SE PUEDE MODIFICAR")
                        'TextBox6.Enabled = False
                        'TextBox3.Enabled = False
                        'Button1.Enabled = False
                        'ComboBox1.Enabled = True
                        'ComboBox2.Enabled = True
                        'ComboBox3.Enabled = True
                    Else
                        If Trim(DataGridView1.Rows(0).Cells(46).Value) = "PESADO" Then
                            TextBox6.Enabled = False
                            TextBox3.Enabled = False
                            Button1.Enabled = True
                            ComboBox1.Enabled = True
                            ComboBox2.Enabled = True
                            ComboBox3.Enabled = True
                            respuesta = MessageBox.Show("EL ROLLO ESTA PESADO QUIERES EDITARLO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                                TextBox6.Enabled = True
                                TextBox3.Enabled = True
                                Button1.Enabled = True
                                ComboBox1.Enabled = True
                                ComboBox2.Enabled = True
                                ComboBox3.Enabled = True
                                TextBox3.Select()
                            Else

                            End If
                        End If
                    End If
                End If



            End If


        End If

    End Sub


    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            Det_Rollo.Label1.Text = "TEJEDORA"
            Det_Rollo.Label2.Text = 1
            Det_Rollo.Show()
        End If
    End Sub
    Dim Rs1 As SqlDataReader
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.F1 Then
            Det_Rollo.Label1.Text = "TRABAJADOR"
            Det_Rollo.Label2.Text = 2
            Det_Rollo.Show()
        Else
            If e.KeyCode = Keys.Enter Then
                abrir()
                Dim sql1 As String = "Select trabajador  FROM tejedores where dni = '" + TextBox3.Text + "' "
                Dim cmd1 As New SqlCommand(sql1, conx)
                Rs1 = cmd1.ExecuteReader()
                If Rs1.Read() = True Then
                    TextBox4.Text = Rs1(0)
                Else
                    MsgBox("SU DNI NO ESTA AGREGADO COMO TEJEDOR")
                    TextBox3.Text = ""
                End If

            End If
            End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "EXTERNO" And TextBox5.Text = "" Then
            MsgBox("FALTA AGREGAR LA GUIA DE REMISION")
        Else
            If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox6.Text = "" Then
                MsgBox("INGRESE INFORMACION DE LOS CAMPOS OBLIGATORIOS")
            Else
                Dim jk As New vpesadorollo
                Dim ll As New ftejeduria
                jk.gccia_3r = Label28.Text
                jk.gprvtej_3r = Trim(TextBox25.Text)
                jk.grollo_3r = TextBox1.Text
                jk.grollom_3r = ""
                jk.gncom_3r = "PESADO  "
                jk.gpedido_3r = TextBox13.Text
                jk.ggalga_3r = TextBox19.Text
                jk.glote_3r = TextBox9.Text
                jk.glinea_3r = TextBox7.Text
                jk.gmaquina_3r = Trim(TextBox2.Text)
                jk.gcantk_3r = Trim(TextBox12.Text)
                jk.gprvhilo_3r = Trim(TextBox24.Text)
                jk.gfcom_3r = DateTimePicker2.Value
                jk.gfproduc_3r = DateTimePicker1.Value
                jk.gpartida_3r = Trim(TextBox14.Text)
                If ComboBox1.Text = "DIURNO" Then
                    jk.gturno_3r = 1
                Else
                    If ComboBox1.Text = "NOCTURNO" Then
                        jk.gturno_3r = 2
                    Else
                        jk.gturno_3r = 3
                    End If

                End If

                jk.ggrm_3r = Trim(TextBox5.Text)
                jk.gsitua_3r = 0
                If Trim(DataGridView1.Rows(0).Cells(46).Value) = "AUDITADO" Or Trim(DataGridView1.Rows(0).Cells(46).Value) = "ACTIVO" Then
                    jk.gflag_3r = 3
                Else
                    jk.gflag_3r = 2
                End If

                jk.gpedmov_3r = Trim(ComboBox2.Text)
                jk.gcantkmv_3r = TextBox6.Text
                jk.gancho_3r = Trim(TextBox22.Text)
                jk.gdiam_3r = Trim(TextBox23.Text)
                jk.gusokar_3r = 0
                jk.gordens_3r = Trim(ComboBox3.Text)
                jk.gccolor_3r = Trim(TextBox11.Text)
                jk.gcalidad_3r = ""
                jk.gubica_3r = ""
                jk.gordtej_3r = Trim(TextBox10.Text)
                jk.gfecgen_3r = DateTimePicker3.Value
                jk.gtejedor_3r = Trim(TextBox3.Text)
                jk.gtiprol_3r = "01"
                jk.gsolmue_3r = ""
                jk.gprvtin_3r = Trim(TextBox25.Text)
                jk.galmacen_3r = ""
                jk.gperiodo_3r = ""
                jk.gvoucher_3r = ""
                jk.gzarmado_3r = 0
                jk.gletradiv_3r = ""
                jk.gmtsafec_3a = "0.00"
                jk.gkgsafec_3a = "0.00"
                jk.gcalrol_3r = ""
                jk.grendimiento_3r = "0.00"
                jk.gobscalidad_3r = ""
                jk.gretrol_3r = ""
                jk.gescalrol_3r = ""
                jk.gdensidad_3r = "0.00"
                jk.gaudicrudo_3r = ""
                jk.gfrevcrudo_3r = DateTimePicker1.Value
                jk.ganchor_3r = "0.00"
                jk.gdensidadr_3r = "0.00"
                jk.grendimientor_3r = "0.00"

                If ll.insertar_pesado_rollos(jk) Then
                    MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
                    TextBox1.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    TextBox7.Text = ""
                    TextBox8.Text = ""
                    TextBox9.Text = ""
                    TextBox10.Text = ""
                    TextBox11.Text = ""
                    TextBox12.Text = ""
                    TextBox13.Text = ""
                    TextBox15.Text = ""
                    TextBox14.Text = ""
                    TextBox17.Text = ""
                    TextBox18.Text = ""
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    'ComboBox2.Items.Clear()
                    'ComboBox3.Items.Clear()
                End If
            End If
        End If


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'If ComboBox1.Text = "EXTERNO" Then
        '    TextBox5.ReadOnly = False
        'Else
        '    If Trim(ComboBox1.Text) = "DIURNO" Then
        '        If Hour(DateTimePicker4.Value) >= "07" And Hour(DateTimePicker4.Value) <= "19" Then
        '        Else
        '            ComboBox1.SelectedIndex = 1
        '            MsgBox("LA HORA CORRESPONDE AL TURNO NOCHE")
        '        End If
        '    Else
        '        If Trim(ComboBox1.Text) = "NOCTURNO" Then
        '            If Hour(DateTimePicker4.Value) > "19" And Hour(DateTimePicker4.Value) < "07" Then

        '            Else
        '                'ComboBox1.SelectedIndex = 0
        '                MsgBox("LA HORA CORRESPONDE AL TURNO DIA")
        '            End If
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox6_LostFocus(sender As Object, e As EventArgs) Handles TextBox6.LostFocus
        Dim NumAuxiliar As Double
        NumAuxiliar = TextBox6.Text

        TextBox6.Text = FormatNumber(NumAuxiliar, 2)
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        Try
            NumConFrac(TextBox6, e)
        Catch ex As Exception
            MsgBox("FALTA INGRESAR UN NUMERO")
        End Try

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

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox3.Enabled = True
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox15.Text = ""
        TextBox14.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        Label33.Visible = False
    End Sub

    Private Sub pesaod_rollo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click

    End Sub

    Private Sub TextBox1_MouseCaptureChanged(sender As Object, e As EventArgs) Handles TextBox1.MouseCaptureChanged

    End Sub

    Private Sub TextBox1_Move(sender As Object, e As EventArgs) Handles TextBox1.Move

    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox1.Select()
        End If
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.Select()
        End If
    End Sub

    Private Sub TextBox1_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles TextBox1.PreviewKeyDown

    End Sub
End Class