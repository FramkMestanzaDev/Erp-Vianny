Imports System.ComponentModel
Imports System.Data.SqlClient
Public Class Form_Clientes
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            MsgBox("FALTA INGRESAR INFORMACION DEL CLIENTE")
        Else
            DESPACHO.Show()
            DESPACHO.TextBox1.Text = TextBox2.Text
            DESPACHO.TextBox3.Text = TextBox1.Text
        End If

    End Sub
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Dim da As New DataTable
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
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub
    Sub llenar_combo_box3()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE rpm_ven ='2' ", conx)
            conn.Fill(ty2)
            ComboBox4.DataSource = ty2
            ComboBox4.DisplayMember = "alias_ven"
            ComboBox4.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Form_Clientes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box3()
        Dim func As New fcliente
        Me.TextBox1.Text = func.CORRELATIVO_CLIENTE() + 1
        Select Case TextBox1.Text.Length


            Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
            Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
            Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
            Case "4" : TextBox1.Text = TextBox1.Text
        End Select
        TextBox21.Text = "0.00"
        ComboBox2.SelectedIndex = 0
        RadioButton2.Checked = True
        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(90)
        blouear()
        TextBox15.Select()
    End Sub
    Sub blouear()
        TextBox12.Enabled = False
        TextBox13.Enabled = False
        TextBox14.Enabled = False
        TextBox2.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox16.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        TextBox11.Enabled = False
        ComboBox3.Enabled = False
        ComboBox2.Enabled = False
        TextBox3.Enabled = False
    End Sub
    Sub desblouear()


        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox16.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        TextBox10.Enabled = True
        TextBox11.Enabled = True
        ComboBox3.Enabled = True
        TextBox3.Enabled = True
        ComboBox2.Enabled = True
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged


    End Sub
    Dim Rsr2, Rsr22 As SqlDataReader
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim c As New vcliente
        Dim R As New vcliente
        Dim c1 As New fcliente
        If Me.ValidateChildren() And ComboBox1.Text = String.Empty Or TextBox4.Text = String.Empty Or TextBox3.Text = String.Empty Or TextBox2.Text = String.Empty Or TextBox11.Text = String.Empty Or TextBox9.Text = String.Empty Or TextBox10.Text = String.Empty Then
            MsgBox("FALTAN INGREAR ALGUNO CAMPOS OBLIGATORIOS")
        Else


            c.gcodigo = TextBox1.Text
            If RadioButton1.Checked = True Then
                c.gT_persona = 1
            Else
                c.gT_persona = 2
            End If
            If ComboBox3.Text = "DNI" Then
                c.gr_social = Trim(TextBox12.Text) & " " & Trim(TextBox13.Text) & " " & Trim(TextBox14.Text)
            Else
                c.gr_social = TextBox2.Text
            End If

            c.gruc = TextBox3.Text
            c.gD_fiscal = TextBox4.Text
            c.gemail = TextBox5.Text
            c.gtelefono = TextBox6.Text
            c.gemail_fac = TextBox7.Text

            'abrir()
            'Dim sql102 As String = "SELECT codigo_ven FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + TextBox17.Text + "'"
            'Dim cmd102 As New SqlCommand(sql102, conx)
            'Rsr2 = cmd102.ExecuteReader()
            'Rsr2.Read()
            c.gVendedor = ComboBox4.SelectedValue.ToString
            'Rsr2.Close()

            'Select Case TextBox17.Text
            '    Case "VPIZARRO" : c.gVendedor = "0027"
            '    Case "VSILVERIO" : c.gVendedor = "0005"
            '    Case "GBALVIN" : c.gVendedor = "0028"
            '    Case "GBEDON" : c.gVendedor = "0010"
            '    Case "VINCIO" : c.gVendedor = "0022"
            '    Case "DBRAVO" : c.gVendedor = "0023"
            '    Case "AMENDO" : c.gVendedor = "0026"
            '    Case "GCUEVA" : c.gVendedor = "0029"
            '    Case "JSALINAS" : c.gVendedor = "0025"
            '    Case "VGRAUS" : c.gVendedor = "0007"
            '    Case "WSALINAS" : c.gVendedor = "0034"
            '    Case "MGRAUS" : c.gVendedor = "0037"
            'End Select

            c.gorigen = ComboBox2.Text
            c.gpais = TextBox8.Text
            c.gdepartamento = TextBox9.Text
            c.gprovincia = TextBox10.Text
            c.gdistrito = TextBox11.Text
            c.gnombres = TextBox12.Text
            c.gapaterno = TextBox13.Text
            c.gamaterno = TextBox14.Text
            c.gCOD_CLI = TextBox15.Text
            If CheckBox1.Checked = True Then
                c.gestado = 1 ' bloqueado
            Else
                c.gestado = 2 'no bloqueado
            End If

            c.glcredito = TextBox21.Text
            Select Case ComboBox3.Text
                Case "DNI" : c.gt_doc = "1"
                Case "RUC" : c.gt_doc = "2"
                Case "CE" : c.gt_doc = "3"
                Case "OTR" : c.gt_doc = "4"

            End Select
            c.gcontacto = TextBox16.Text
            c.gfcom = DateTimePicker1.Value
            c.gfcom_fin = DateTimePicker2.Value
            Select Case ComboBox1.Text


                Case "PRENDA" : c.gU_NEGOCIO = "1"
                Case "TELA PLANA" : c.gU_NEGOCIO = "2"
                Case "TELA PUNTO" : c.gU_NEGOCIO = "3"
                Case "TELA NOTEX" : c.gU_NEGOCIO = "4"
                Case "INDUMENTARIA MEDICA" : c.gU_NEGOCIO = "5"
                Case "HILO" : c.gU_NEGOCIO = "6"
            End Select
            R.gcodigo = TextBox1.Text
            c1.eliminar_CLIENTE(R)
            c1.insertar_cliente(c)
            MsgBox("SE REGISTRO AL NUEVO CLIENTE")

            Dim func As New fcliente
            Me.TextBox1.Text = func.CORRELATIVO_CLIENTE() + 1
            Select Case TextBox1.Text.Length


                Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = TextBox1.Text
            End Select

            MsgBox(DateTimePicker1.Value)
            MsgBox(DateTimePicker2.Value)
            'TextBox18.Text = 3

            If TextBox18.Text = 2 Then

                Dim c14 As New vclientevianny
                Dim c15 As New fcliente

                c14.gcodigo = TextBox15.Text
                c14.gnombre = TextBox2.Text
                If ComboBox3.Text = "DNI" Then
                    c14.gnombre = Trim(TextBox12.Text) & " " & Trim(TextBox13.Text) & " " & Trim(TextBox14.Text)
                Else
                    c14.gnombre = TextBox2.Text
                End If

                If RadioButton1.Checked = True Then
                    c14.gtdoc = "01"
                Else
                    If RadioButton2.Checked = True Then
                        c14.gtdoc = "06"
                    End If

                End If

                c14.gruc = TextBox3.Text
                c14.gdireccion = TextBox4.Text
                Select Case ComboBox2.Text
                    Case "NACIONAL" : c14.gorigen = "N"
                    Case "EXTRANJERO" : c14.gorigen = "E"
                End Select

                c14.gpais = "9589"
                c14.gemail_fac = TextBox7.Text
                c14.gubigeo = TextBox19.Text
                Select Case ComboBox3.Text
                    Case "RUC" : c14.gtiper = "2"
                    Case "DNI" : c14.gtiper = "1"
                    Case "OTR" : c14.gtiper = "2"
                End Select

                c14.gnomf = TextBox12.Text
                c14.gamaterno = TextBox14.Text
                c14.gapaterno = TextBox13.Text

                c15.insertar_cliente_vianny(c14)
                TextBox15.Text = ""
                LIMPIAR()
            End If
            TextBox2.Select()
            ComboBox2.SelectedIndex = 0
            blouear()
            LIMPIAR()
        End If
    End Sub

    Sub LIMPIAR()
        TextBox2.Text = ""
        TextBox4.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""

        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox14.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Dim hj, hj2, HJ3 As New DataTable

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Ubigeo.Show()
    End Sub

    Private Sub Form_Clientes_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.Text = "DNI" Then
            TextBox3.Text = Mid(TextBox3.Text, 1, 8)
            RadioButton1.Checked = True
            RadioButton2.Checked = False
            TextBox12.Enabled = True
            TextBox13.Enabled = True
            TextBox14.Enabled = True
            TextBox2.Enabled = False
        Else
            ComboBox3.Text = "RUC"
            RadioButton1.Checked = False
            RadioButton2.Checked = True
            TextBox12.Enabled = False
            TextBox13.Enabled = False
            TextBox14.Enabled = False
            TextBox2.Enabled = True
        End If
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.DataSource = ""
            DataGridView2.DataSource = ""
            DataGridView3.DataSource = ""
            'TextBox12.Enabled = False
            'TextBox13.Enabled = False
            'TextBox14.Enabled = False
            'TextBox2.Enabled = True
            ComboBox3.SelectedIndex = 1
            Select Case TextBox15.Text.Length

                Case "1" : TextBox15.Text = TextBox15.Text & "" & "0000000000"
                Case "2" : TextBox15.Text = TextBox15.Text & "" & "000000000"
                Case "3" : TextBox15.Text = TextBox15.Text & "" & "00000000"
                Case "4" : TextBox15.Text = TextBox15.Text & "" & "0000000"
                Case "5" : TextBox15.Text = TextBox15.Text & "" & "000000"
                Case "6" : TextBox15.Text = TextBox15.Text & "" & "00000"
                Case "7" : TextBox15.Text = TextBox15.Text & "" & "0000"
                Case "8" : TextBox15.Text = TextBox15.Text & "" & "000"
                Case "9" : TextBox15.Text = TextBox15.Text & "" & "00"
                Case "10" : TextBox15.Text = TextBox15.Text & "" & "0"
                Case "11" : TextBox15.Text = TextBox15.Text
            End Select



            Dim vj As New vcliente
            Dim jk As New vcliente
            Dim K As New vcliente
            Dim vj1 As New fcliente

            K.gruc = TextBox15.Text
            HJ3 = vj1.CLIENTE_FRANCISCO(K)
            If HJ3.Rows.Count <> 0 Then
                DateTimePicker2.Value = DateTimePicker1.Value.AddDays(90)
                DataGridView3.DataSource = HJ3
                Dim vend As String
                vend = ""
                abrir()
                Dim sql1021 As String = "SELECT alias_ven FROM  custom_vianny.dbo.Vendedores WHERE codigo_ven ='" + DataGridView3.Rows(0).Cells(12).Value + "'"
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr22 = cmd1021.ExecuteReader()
                Rsr22.Read()
                vend = Rsr22(0)
                Rsr22.Close()
                'Select Case DataGridView3.Rows(0).Cells(12).Value
                '    Case "0027" : vend = "VPIZARRO"
                '    Case "0005" : vend = "VSILVERIO"
                '    Case "0028" : vend = "GBALVIN"
                '    Case "0010" : vend = "GBEDON"
                '    Case "0022" : vend = "VINCIO"
                '    Case "0023" : vend = "DBRAVO"
                '    Case "0026" : vend = "AMENDO"
                '    Case "0029" : vend = "GCUEVA"
                '    Case "0025" : vend = "JSALINAS"
                '    Case "0007" : vend = "VGRAUS"
                '    Case "0030" : vend = "RMEDINA"
                '    Case "0034" : vend = "WSALINAS"
                'End Select


                'If Trim(MDIParent1.Label3.Text) <> Trim(vend) Then
                '    MsgBox("EL CLIENTE ESTA ASIGNADO A OTRO VENDEDOR")

                'Else

                TextBox1.Text = DataGridView3.Rows(0).Cells(0).Value
                If DataGridView3.Rows(0).Cells(1).Value = 1 Then
                    RadioButton1.Checked = True
                Else
                    If DataGridView3.Rows(0).Cells(1).Value = 2 Then
                        RadioButton2.Checked = True
                    End If
                End If
                TextBox12.Text = DataGridView3.Rows(0).Cells(2).Value
                TextBox13.Text = DataGridView3.Rows(0).Cells(3).Value
                TextBox14.Text = DataGridView3.Rows(0).Cells(4).Value
                TextBox2.Text = DataGridView3.Rows(0).Cells(5).Value
                TextBox3.Text = DataGridView3.Rows(0).Cells(6).Value

                'Select Case DataGridView3.Rows(0).Cells(7).Value
                '    Case "01" : ComboBox3.Text = "DNI"
                '    Case "06" : ComboBox3.Text = "RUC"

                '    Case "20" : ComboBox3.Text = "OTR"

                'End Select
                TextBox4.Text = DataGridView3.Rows(0).Cells(8).Value
                TextBox5.Text = DataGridView3.Rows(0).Cells(9).Value
                TextBox6.Text = DataGridView3.Rows(0).Cells(10).Value
                TextBox7.Text = DataGridView3.Rows(0).Cells(11).Value

                'Select Case DataGridView3.Rows(0).Cells(12).Value
                '    Case "0027" : TextBox17.Text = "VPIZARRO"
                '    Case "0005" : TextBox17.Text = "VSILVERIO"
                '    Case "0028" : TextBox17.Text = "GBALVIN"
                '    Case "0010" : TextBox17.Text = "GBEDON"
                '    Case "0022" : TextBox17.Text = "VINCIO"
                '    Case "0023" : TextBox17.Text = "DBRAVO"
                '    Case "0026" : TextBox17.Text = "AMENDO"
                '    Case "0029" : TextBox17.Text = "GCUEVA"
                '    Case "0025" : TextBox17.Text = "JSALINAS"
                '    Case "0007" : TextBox17.Text = "VGRAUS"
                '    Case "0030" : TextBox17.Text = "RMEDINA"
                '    Case "0034" : TextBox17.Text = "WSALINAS"
                'End Select
                TextBox16.Text = DataGridView3.Rows(0).Cells(13).Value
                ComboBox2.Text = DataGridView3.Rows(0).Cells(14).Value
                TextBox9.Text = DataGridView3.Rows(0).Cells(15).Value
                TextBox9.Text = DataGridView3.Rows(0).Cells(16).Value
                TextBox10.Text = DataGridView3.Rows(0).Cells(17).Value
                TextBox11.Text = DataGridView3.Rows(0).Cells(18).Value
                TextBox21.Text = DataGridView3.Rows(0).Cells(21).Value

                Select Case DataGridView3.Rows(0).Cells(22).Value
                    Case "1" : ComboBox1.Text = "PRENDA"
                    Case "2" : ComboBox1.Text = "TELA PLANA"
                    Case "3" : ComboBox1.Text = "TELA PUNTO"
                    Case "4" : ComboBox1.Text = "TELA NOTEX"
                    Case "5" : ComboBox1.Text = "INDUMENTARIA MEDICA"
                    Case "6" : ComboBox1.Text = "HILO"
                End Select
                Select Case DataGridView3.Rows(0).Cells(7).Value
                    Case "1" : ComboBox3.Text = "DNI"
                    Case "2" : ComboBox3.Text = "RUC"
                    Case "3" : ComboBox3.Text = "CE"
                    Case "4" : ComboBox3.Text = "OTR"

                End Select
                If DataGridView3.Rows(0).Cells(20).Value = 1 Then
                    CheckBox1.Checked = True
                Else
                    CheckBox1.Checked = False
                End If
                desblouear()
                'End If

                TextBox18.Text = 1
                TextBox20.Text = 1
            Else
                Dim func As New fcliente
                Me.TextBox1.Text = func.CORRELATIVO_CLIENTE() + 1
                Select Case TextBox1.Text.Length


                    Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
                    Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
                    Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
                    Case "4" : TextBox1.Text = TextBox1.Text
                End Select

                vj.gruc = TextBox15.Text
                hj = vj1.buscar_cliente_vianny(vj)
                If hj.Rows.Count <> 0 Then
                    TextBox20.Text = 2
                    DataGridView1.DataSource = hj
                    Dim j As Integer
                    j = DataGridView1.Rows.Count

                    For i = 0 To j - 1
                        If DataGridView1.Rows(i).Cells(8).Value = "06" Then
                            RadioButton2.Checked = True
                        Else
                            RadioButton1.Checked = True
                        End If


                        Select Case DataGridView1.Rows(i).Cells(8).Value
                            Case "01" : ComboBox3.Text = "DNI"
                            Case "06" : ComboBox3.Text = "RUC"

                            Case "20" : ComboBox3.Text = "OTR"

                        End Select
                        jk.gubigeo = DataGridView1.Rows(i).Cells(4).Value
                        hj2 = vj1.buscar_direccion(jk)
                        If hj2.Rows.Count <> 0 Then
                            DataGridView2.DataSource = hj2
                            TextBox9.Text = DataGridView2.Rows(0).Cells(2).Value
                            TextBox10.Text = DataGridView2.Rows(0).Cells(1).Value
                            TextBox11.Text = DataGridView2.Rows(0).Cells(0).Value
                        End If

                        TextBox12.Text = DataGridView1.Rows(i).Cells(5).Value
                        TextBox13.Text = DataGridView1.Rows(i).Cells(6).Value
                        TextBox14.Text = DataGridView1.Rows(i).Cells(7).Value
                        TextBox2.Text = DataGridView1.Rows(i).Cells(1).Value
                        TextBox4.Text = DataGridView1.Rows(i).Cells(2).Value
                        TextBox3.Text = DataGridView1.Rows(i).Cells(0).Value
                    Next
                    TextBox18.Text = 1
                    desblouear()
                Else
                    TextBox18.Text = 2
                    desblouear()
                    Button2.Enabled = True
                    TextBox2.Enabled = True
                    TextBox12.Text = ""
                    TextBox13.Text = ""
                    TextBox14.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = TextBox15.Text
                    TextBox4.Text = ""
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    TextBox7.Text = ""
                    TextBox21.Text = "0.00"
                    ComboBox1.SelectedIndex = 0
                End If
            End If


        End If

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click
        TextBox12.Enabled = True
        TextBox13.Enabled = True
        TextBox14.Enabled = True
        TextBox2.Enabled = False
        ComboBox3.SelectedIndex = 0
        TextBox3.Text = Microsoft.VisualBasic.Left(TextBox15.Text, 8)
        desblouear()
    End Sub

    Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
        TextBox12.Enabled = False
        TextBox13.Enabled = False
        TextBox14.Enabled = False
        TextBox2.Enabled = True
        ComboBox3.SelectedIndex = 1
        TextBox3.Text = TextBox15.Text
        desblouear()
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        TextBox2.Text = TextBox13.Text + " " + TextBox14.Text + "," + TextBox12.Text
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        TextBox2.Text = TextBox13.Text + " " + TextBox14.Text + "," + TextBox12.Text
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        TextBox2.Text = TextBox13.Text + " " + TextBox14.Text + "," + TextBox12.Text
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_Validating(sender As Object, e As CancelEventArgs) Handles TextBox3.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL RUC O DNI")
        End If
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(90)
    End Sub

    Private Sub TextBox4_Validating(sender As Object, e As CancelEventArgs) Handles TextBox4.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR DIRECCION FISCAL")
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub

    Private Sub TextBox2_Validating(sender As Object, e As CancelEventArgs) Handles TextBox2.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL CLIENTE")
        End If
    End Sub

    Private Sub TextBox9_Validating(sender As Object, e As CancelEventArgs) Handles TextBox9.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL DEPARTAMENTO")
        End If
    End Sub

    Private Sub TextBox10_Validating(sender As Object, e As CancelEventArgs) Handles TextBox10.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR LA PROVINCIA")
        End If
    End Sub

    Private Sub TextBox11_Validating(sender As Object, e As CancelEventArgs) Handles TextBox11.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL DISTRITO  ")
        End If
    End Sub

    Private Sub TextBox15_QueryAccessibilityHelp(sender As Object, e As QueryAccessibilityHelpEventArgs) Handles TextBox15.QueryAccessibilityHelp

    End Sub
End Class