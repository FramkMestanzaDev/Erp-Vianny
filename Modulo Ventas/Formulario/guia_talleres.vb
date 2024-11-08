Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class guia_talleres
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado1 As SqlCommand
    Public respuesta1 As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Public comando As SqlCommand
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

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
    Sub correlativo()
        abrir()

        Dim sql1021 As String = "SELECT TOP 1 substring(ncom_3,5,8) as ncom_3 FROM custom_vianny.dbo.Qag3000 WHERE Ano_3 ='" + Label29.Text + "' AND talm_3 = '" + Label23.Text + "' and Empr_3 ='" + Trim(Label30.Text) + "' and ccom_3 ='2' and ncom_3 like '14%' ORDER BY ncom_3 desc"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr21 = cmd1021.ExecuteReader()
        If Rsr21.Read() = True Then

            Label3.Text = Rsr21(0) + 1
            Select Case Label3.Text.Length
                Case "1" : Label3.Text = "14" & "00000" & Label3.Text
                Case "2" : Label3.Text = "14" & "0000" & Label3.Text
                Case "3" : Label3.Text = "14" & "000" & Label3.Text
                Case "4" : Label3.Text = "14" & "00" & Label3.Text
                Case "5" : Label3.Text = "14" & "0" & Label3.Text
                Case "6" : Label3.Text = "14" & Label3.Text
            End Select
        Else

            Label3.Text = 1
            Select Case Label3.Text.Length
                Case "1" : Label3.Text = "14" & "00000" & Label3.Text
                Case "2" : Label3.Text = "14" & "0000" & Label3.Text
                Case "3" : Label3.Text = "14" & "000" & Label3.Text
                Case "4" : Label3.Text = "14" & "00" & Label3.Text
                Case "5" : Label3.Text = "14" & "0" & Label3.Text
                Case "6" : Label3.Text = "14" & Label3.Text
            End Select
        End If
        Rsr21.Close()
    End Sub
    Private Sub guia_talleres_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim func As New fguiasistema
        Dim fu As New vguiacabecera
        Dim fu1 As New vguiacabecera
        'If Label30.Text = "01" Then
        '    TextBox1.ReadOnly = True
        'Else
        '    TextBox1.ReadOnly = False
        'End If
        fu.gserie = TextBox1.Text
        fu.gccia = Label30.Text
        TextBox2.Text = func.mostrar_guia_correlativo(fu) + 1
        Select Case TextBox2.Text.Length

            Case "1" : TextBox2.Text = "0000000" & "" & TextBox2.Text
            Case "2" : TextBox2.Text = "000000" & "" & TextBox2.Text
            Case "3" : TextBox2.Text = "00000" & "" & TextBox2.Text
            Case "4" : TextBox2.Text = "0000" & "" & TextBox2.Text
            Case "5" : TextBox2.Text = "000" & "" & TextBox2.Text
            Case "6" : TextBox2.Text = "00" & "" & TextBox2.Text
            Case "7" : TextBox2.Text = "0" & "" & TextBox2.Text
            Case "8" : TextBox2.Text = TextBox2.Text
        End Select

        'fu1.gmes = DateTime.Now.ToString("MM")
        'fu1.galmacen = Label23.Text
        'fu1.gano = Label29.Text
        'fu1.gccia = Label30.Text
        'Label3.Text = func.mostrar_correlativo_nsa(fu1) + 1
        'Select Case Label3.Text.Length

        '    Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
        '    Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
        '    Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
        '    Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
        'End Select
        abrir()
        llenar_combo_box()
        'llenar_combo_box2(ComboBox2, Label23.Text)
        TextBox2.Select()
        RadioButton1.Checked = True
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox5.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        TextBox11.Enabled = False
        TextBox13.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False


        Button5.Enabled = False
        correlativo()
        'ComboBox1.SelectedIndex = 0
        ''ComboBox2.SelectedIndex = 0
        'ComboBox3.SelectedIndex = 0
    End Sub
    Dim dt1, dt2 As New DataTable

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Dim hg As DataTable

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        Dim i As Integer
        Dim ml As New vguiacabecera
        Dim ml1 As New vguiacabecera
        Dim mk As New fguiasistema
        i = Len(TextBox2.Text)
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        If e.KeyCode = Keys.Enter Then
            If Label33.Text = 0 Then
                DataGridView1.Rows.Clear()
                'TextBox22.Text = 0
                'TextBox21.Text = 0
                Select Case i
                    Case "1" : TextBox2.Text = "0000000" & "" & TextBox2.Text
                    Case "2" : TextBox2.Text = "000000" & "" & TextBox2.Text
                    Case "3" : TextBox2.Text = "00000" & "" & TextBox2.Text
                    Case "4" : TextBox2.Text = "0000" & "" & TextBox2.Text
                    Case "5" : TextBox2.Text = "000" & "" & TextBox2.Text
                    Case "6" : TextBox2.Text = "00" & "" & TextBox2.Text
                    Case "7" : TextBox2.Text = "0" & "" & TextBox2.Text
                    Case "8" : TextBox2.Text = TextBox2.Text
                End Select
                ml.gcorrelativo = TextBox2.Text
                ml.gserie = TextBox1.Text
                ml.galmacen = Label23.Text
                ml.gccia = Label30.Text
                ml1.gcorrelativo = TextBox2.Text
                ml1.gserie = TextBox1.Text
                ml1.gccia = Label30.Text
                If mk.mostrar_correlativo_guia1(ml) = True Then
                    If mk.mostrar_almacen_guia_prod(ml) = Label23.Text Then
                        Dim jk As New fguiasistema

                        dt1 = jk.mostrar_cabecera_guia_prod(ml)
                        dt2 = jk.mostrar_detalle_guia(ml1)

                        If dt1.Rows.Count <> 0 Then
                            DataGridView2.DataSource = dt1
                            TextBox8.Text = DataGridView2.Rows(0).Cells(0).Value
                            TextBox9.Text = DataGridView2.Rows(0).Cells(1).Value
                            TextBox10.Text = DataGridView2.Rows(0).Cells(2).Value
                            TextBox16.Text = DataGridView2.Rows(0).Cells(3).Value
                            TextBox14.Text = DataGridView2.Rows(0).Cells(4).Value
                            TextBox15.Text = DataGridView2.Rows(0).Cells(5).Value
                            TextBox12.Text = DataGridView2.Rows(0).Cells(9).Value
                            TextBox13.Text = DataGridView2.Rows(0).Cells(10).Value
                            TextBox11.Text = DataGridView2.Rows(0).Cells(13).Value
                            TextBox4.Text = DataGridView2.Rows(0).Cells(14).Value
                            TextBox6.Text = DataGridView2.Rows(0).Cells(18).Value
                            TextBox21.Text = DataGridView2.Rows(0).Cells(16).Value
                            TextBox22.Text = DataGridView2.Rows(0).Cells(20).Value
                            TextBox18.Text = "FAC"
                            TextBox7.Text = DataGridView2.Rows(0).Cells(6).Value
                            TextBox17.Text = DataGridView2.Rows(0).Cells(7).Value
                            Label3.Text = DataGridView2.Rows(0).Cells(11).Value & "" & DataGridView2.Rows(0).Cells(12).Value
                            TextBox20.Text = DataGridView2.Rows(0).Cells(19).Value
                            DateTimePicker1.Value = DataGridView2.Rows(0).Cells(21).Value
                            DateTimePicker2.Value = DataGridView2.Rows(0).Cells(22).Value
                            enunciado1 = New SqlCommand("select rtrim(ltrim(nomb_17f)) AS nomb_17f from custom_vianny.dbo.Fag1700 where ccia_17F ='" + Trim(Label30.Text) + "' and motiv_17f = '" + Trim(DataGridView2.Rows(0).Cells(15).Value) + "'", conx)
                            respuesta1 = enunciado1.ExecuteReader
                            While respuesta1.Read
                                ComboBox1.Text = respuesta1.Item("nomb_17f")
                            End While
                            respuesta1.Close()

                            If DataGridView2.Rows(0).Cells(8).Value = 0 Then
                                RadioButton2.Checked = True
                                RadioButton1.Checked = False
                                RadioButton1.Enabled = False
                                Label28.Visible = True
                            Else
                                RadioButton2.Checked = False
                                RadioButton1.Checked = True
                                RadioButton1.Enabled = False
                                Label28.Visible = False
                            End If
                            If DataGridView2.Rows(0).Cells(17).Value = "EXT" Then
                                ComboBox2.SelectedIndex = 0
                            Else
                                ComboBox2.SelectedIndex = 1
                            End If

                        End If
                        If dt2.Rows.Count <> 0 Then
                            DataGridView3.DataSource = dt2
                            Dim j As Integer
                            j = DataGridView3.Rows.Count
                            DataGridView1.Rows.Add(j)
                            For p = 0 To j - 1
                                DataGridView1.Rows(p).Cells(0).Value = DataGridView3.Rows(p).Cells(4).Value
                                DataGridView1.Rows(p).Cells(1).Value = DataGridView3.Rows(p).Cells(6).Value
                                DataGridView1.Rows(p).Cells(2).Value = DataGridView3.Rows(p).Cells(5).Value
                                DataGridView1.Rows(p).Cells(3).Value = DataGridView3.Rows(p).Cells(3).Value
                                DataGridView1.Rows(p).Cells(4).Value = DataGridView3.Rows(p).Cells(0).Value
                                DataGridView1.Rows(p).Cells(5).Value = DataGridView3.Rows(p).Cells(8).Value
                                DataGridView1.Rows(p).Cells(6).Value = DataGridView3.Rows(p).Cells(1).Value
                                DataGridView1.Rows(p).Cells(7).Value = DataGridView3.Rows(p).Cells(2).Value
                                DataGridView1.Rows(p).Cells(8).Value = DataGridView3.Rows(p).Cells(7).Value
                                DataGridView1.Rows(p).Cells(11).Value = DataGridView3.Rows(p).Cells(10).Value
                                'DataGridView1.Rows(p).Cells(7).Value = DataGridView3.Rows(p).Cells(9).Value
                                If Label23.Text = "0400" Or Label23.Text = "0700" Then
                                    abrir()
                                    llenar_combo_box4(Trim(DataGridView3.Rows(p).Cells(6).Value))

                                    If Trim(DataGridView3.Rows(p).Cells(9).Value) = "" Then

                                    Else
                                        DataGridView1.Rows(p).Cells(9).Value = DataGridView3.Rows(p).Cells(9).Value.ToString
                                    End If


                                End If



                            Next
                            DataGridView1.Columns(0).Frozen = True
                            DataGridView1.Columns(1).Frozen = True
                            DataGridView1.Columns(2).Frozen = True

                        End If
                        Button3.Enabled = False
                        Button1.Enabled = False
                        Button2.Enabled = False
                        'noeditable()
                        DataGridView1.Enabled = False

                        'PictureBox2.Enabled = True
                        'PictureBox5.Enabled = True

                        If RadioButton2.Checked = True Then
                            Label15.Visible = True
                        Else
                            Label15.Visible = False
                        End If
                        TextBox21.Enabled = False
                        'RadioButton2.Checked = True
                        'RadioButton1.Checked = True
                        'PictureBox1.Enabled = True
                    Else
                        MsgBox("LA GUIA PERTENECE A OTRO ALMACEN")
                    End If
                Else


                    Select Case TextBox11.Text.Length
                        Case "1" : TextBox11.Text = "0000000" & "" & TextBox11.Text
                        Case "2" : TextBox11.Text = "000000" & "" & TextBox11.Text
                        Case "3" : TextBox11.Text = "00000" & "" & TextBox11.Text
                        Case "4" : TextBox11.Text = "0000" & "" & TextBox11.Text
                        Case "5" : TextBox11.Text = "000" & "" & TextBox11.Text
                        Case "6" : TextBox11.Text = "00" & "" & TextBox11.Text
                        Case "7" : TextBox11.Text = "0" & "" & TextBox11.Text
                        Case "8" : TextBox11.Text = TextBox11.Text
                    End Select
                    correlativo()
                    TextBox3.Enabled = True
                    TextBox4.Enabled = True
                    TextBox8.Enabled = True
                    TextBox10.Enabled = True
                    TextBox11.Enabled = True
                    TextBox13.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    TextBox20.Enabled = True
                    TextBox21.Enabled = True
                    TextBox25.Enabled = True
                    'PictureBox1.Enabled = True

                    Button5.Enabled = True
                    'PictureBox4.Enabled = True


                    TextBox4.Text = ""


                    ComboBox1.Text = ""
                    ComboBox2.Text = ""

                End If
                TextBox4.Select()
            Else
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                TextBox8.Enabled = True
                TextBox10.Enabled = True
                TextBox11.Enabled = True
                TextBox13.Enabled = True
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                TextBox20.Enabled = True
                TextBox21.Enabled = True
                TextBox25.Enabled = True
                Button5.Enabled = True
                Button3.Enabled = True
                Button1.Enabled = True
                Button2.Enabled = True
            End If
        Else

            If e.KeyCode = Keys.F1 Then
                Dim ds As New Det_guia_talleres
                ds.TextBox2.Text = Label30.Text
                ds.Label3.Text = Label23.Text
                ds.ShowDialog()
            End If

        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Sub llenar_combo_box()
        Try

            conn = New SqlDataAdapter("Select rtrim(ltrim(motiv_17f)) As motiv_17f,rtrim(ltrim(nomb_17f)) As nomb_17f from custom_vianny.dbo.Fag1700", conx)
            conn.Fill(ty)
            ComboBox1.DataSource = ty
            ComboBox1.DisplayMember = "nomb_17f"
            ComboBox1.ValueMember = "motiv_17f"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged

    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.F1 Then
            Dim res As New Transportistas
            res.TextBox1.Text = 3
            res.Show()

        Else
            If e.KeyCode = Keys.Enter Then
                abrir()
                enunciado3 = New SqlCommand("Select nomb_18 , direcc_18,ruc_18,chofer_18,placa_18 As PLA,brevete_18 from custom_vianny.dbo.Fag1800 where empr_18 =" + Trim(Label30.Text) + " And trans_18 =" + TextBox4.Text, conx)
                respuesta3 = enunciado3.ExecuteReader
                While respuesta3.Read
                    TextBox5.Text = respuesta3.Item("nomb_18")
                    TextBox12.Text = respuesta3.Item("nomb_18")
                    TextBox13.Text = respuesta3.Item("direcc_18")
                    TextBox11.Text = respuesta3.Item("ruc_18")
                    TextBox16.Text = respuesta3.Item("chofer_18")
                    TextBox14.Text = respuesta3.Item("PLA")
                    TextBox15.Text = respuesta3.Item("brevete_18")

                End While
                respuesta3.Close()
                TextBox8.Select()
            End If

        End If
    End Sub

    Private Sub TextBox21_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox21.KeyDown
        If e.KeyCode = Keys.F1 Then
            det_ns_Prodc.Label2.Text = Label23.Text
            det_ns_Prodc.Label3.Text = 5
            det_ns_Prodc.Label5.Text = Trim(Label30.Text)
            det_ns_Prodc.Show()
        End If
    End Sub
    Dim Rsr2q, Rsr30014, Rsr233 As SqlDataReader
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If RadioButton2.Checked = True Then
            MsgBox("LA GUIA ESTA ANULADA NO SE PUEDE MODIFICAR")
        Else
            abrir()
            Dim va As String
            Dim sql102 As String = "Select cerrado_3 FROM custom_vianny.dbo.fag0400 WHERE sfactu_3 ='" + Trim(TextBox1.Text) + "' AND nfactu_3 ='" + Trim(TextBox2.Text) + "' AND CIA_3 ='" + Trim(Label30.Text) + "'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2q = cmd102.ExecuteReader()

            If Rsr2q.Read() = True Then
                va = Rsr2q(0)

            End If
            Rsr2q.Close()
            Dim sql1022134 As String = "select COUNT( nfactu_3) from custom_vianny.dbo.fag0400 where  CIA_3 =" + Trim(Label30.Text) + " and sfactu_3 ='" + Trim(TextBox1.Text) + "' and nfactu_3 =" + Trim(TextBox2.Text) + "  and cerrado_3 ='2'"
            Dim cmd1022134 As New SqlCommand(sql1022134, conx)
            Rsr30014 = cmd1022134.ExecuteReader()
            Dim jo4 As Integer
            If Rsr30014.Read() Then
                jo4 = Rsr30014(0)
            Else
                jo4 = 0
            End If
            Rsr30014.Close()

            If va = True Then

                MsgBox("LA GUIA DE REMISION ESTA REVISADA POR CONTABILIDAD NO SE PUEDE MODIFICAR")
            Else
                If jo4 = 0 Then
                    sieditable()
                    DataGridView1.Enabled = True
                    Button3.Enabled = True
                    Button1.Enabled = True
                    Button2.Enabled = True
                    Label34.Text = "1"
                Else
                    MsgBox("LA GUIA DE REMISION ESTA ENVIADA A SUNAT NO SE PEUDE MODIFICAR")

                End If

            End If
        End If

    End Sub
    Public Sub sieditable()

        TextBox20.Enabled = True
        TextBox21.Enabled = True
        TextBox25.Enabled = True
        TextBox8.Enabled = True
        TextBox10.Enabled = True
        TextBox11.Enabled = True
        TextBox12.Enabled = True
        TextBox13.Enabled = True
        TextBox14.Enabled = True
        TextBox15.Enabled = True
        TextBox16.Enabled = True
        TextBox4.Enabled = True
        DateTimePicker1.Enabled = True
        DateTimePicker2.Enabled = True
        TextBox21.Enabled = True
        Button5.Enabled = True

        ComboBox1.Enabled = True
        ComboBox2.Enabled = True

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim I18, VAL, TAB As Integer
        TAB = DataGridView1.Rows.Count
        If TAB <> 0 Then
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

                I18 = DataGridView1.Rows.Count


                For i1 = 0 To I18 - 1

                    VAL = i1 + 1
                    Select Case VAL.ToString.Length
                        Case "1" : DataGridView1.Rows(i1).Cells(0).Value = "00" & "" & VAL
                        Case "2" : DataGridView1.Rows(i1).Cells(0).Value = "0" & "" & VAL
                        Case "3" : DataGridView1.Rows(i1).Cells(0).Value = VAL

                    End Select
                Next

            End If
        Else
            MsgBox("NO HAY PRODUCTOS A ELIMINAR")
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
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Me.ValidateChildren() And Trim(TextBox21.Text) = String.Empty Or Trim(TextBox8.Text) = String.Empty Then
            MsgBox("FALTAN INGREAR ALGUNO CAMPOS OBLIGATORIOS")
        Else
            abrir()
            Dim agregar As String = "DELETE FROM custom_vianny.dbo.Fap0400 Where CIA_3a = '" + Trim(Label30.Text) + "' And sfactu_3a = '" + Trim(TextBox1.Text) + "' And nfactu_3a = '" + Trim(TextBox2.Text) + "'"
            Dim agregar11 As String = "DELETE FROM custom_vianny.dbo.Fag0400 Where CIA_3 = '" + Trim(Label30.Text) + "' And sfactu_3 = '" + Trim(TextBox1.Text) + "' And nfactu_3 = '" + Trim(TextBox2.Text) + "' "
            Dim agregar1 As String = "DELETE FROM custom_vianny.dbo.Fat0400 Where Cia_3t = '" + Trim(Label30.Text) + "' And sfactu_3t = '" + Trim(TextBox1.Text) + "' And nfactu_3t = '" + Trim(TextBox2.Text) + "'"
            Dim agregar2 As String = "DELETE FROM custom_vianny.dbo.Fah0400 Where CIA_3h = '" + Trim(Label30.Text) + "' And sfactu_3h = '" + Trim(TextBox1.Text) + "' And nfactu_3h = '" + Trim(TextBox2.Text) + "'"
            Dim agregar3 As String = "DELETE FROM custom_vianny.dbo.Fad0400 Where CIA_3d = '" + Trim(Label30.Text) + "' And sfactu_3d = '" + Trim(TextBox1.Text) + "' And nfactu_3d = '" + Trim(TextBox2.Text) + "'"
            Dim agregar4 As String = "DELETE FROM custom_vianny.dbo.Fat0400 Where Cia_3t = '" + Trim(Label30.Text) + "' And sfactu_3t = '" + Trim(TextBox1.Text) + "' And nfactu_3t = '" + Trim(TextBox2.Text) + "'"
            Dim agregar5 As String = "DELETE FROM custom_vianny.dbo.Faq0400 Where CIA_4q = '" + Trim(Label30.Text) + "' And sfactu_4q = '" + Trim(TextBox1.Text) + "' And nfactu_4q = '" + Trim(TextBox2.Text) + "'"
            Dim agregar6 As String = "DELETE FROM custom_vianny.dbo.Qap3000 Where Empr_3a = '" + Trim(Label30.Text) + "' And Ano_3a = '" + Trim(Label29.Text) + "' And talm_3a = '" + Label23.Text + "' And ccom_3a = '2' And ncom_3a = '" + Label3.Text + "'"
            Dim agregar7 As String = "DELETE FROM custom_vianny.dbo.Qat3000 Where Empr_3t = '" + Trim(Label30.Text) + "' And Ano_3t = '" + Trim(Label29.Text) + "' And talm_3t = '" + Label23.Text + "' And ccom_3t = '2' And ncom_3t = '" + Label3.Text + "'"
            Dim agregar8 As String = "DELETE FROM custom_vianny.dbo.QAD3000 Where Empr_3d = '" + Trim(Label30.Text) + "' And Ano_3d = '" + Trim(Label29.Text) + "' And talm_3d = '" + Label23.Text + "' And ccom_3d = '2' And ncom_3d = '" + Label3.Text + "'"
            Dim agregar9 As String = "DELETE FROM custom_vianny.dbo.QAD3000 Where Empr_3d = '" + Trim(Label30.Text) + "' And Ano_3d = '" + Trim(Label29.Text) + "' And talm_3d = '" + Label23.Text + "' And ccom_3d = '2' And ncom_3d = '" + Label3.Text + "'"
            Dim agregar10 As String = "DELETE FROM custom_vianny.dbo.Qag3000 Where Empr_3 = '" + Trim(Label30.Text) + "' And Ano_3 = '" + Trim(Label29.Text) + "' And talm_3 = '" + Label23.Text + "' And ccom_3 = '2' And ncom_3 = '" + Label3.Text + "'"
            ELIMINAR(agregar)
            ELIMINAR(agregar1)
            ELIMINAR(agregar2)
            ELIMINAR(agregar3)
            ELIMINAR(agregar4)
            ELIMINAR(agregar5)
            ELIMINAR(agregar6)
            ELIMINAR(agregar7)
            ELIMINAR(agregar8)
            ELIMINAR(agregar9)
            ELIMINAR(agregar10)
            ELIMINAR(agregar11)
            Dim o As Integer
            o = DataGridView1.Rows.Count

            Dim cmd13 As New SqlCommand("INSERT INTO custom_vianny.dbo.Fag0400 (CIA_3  , sfactu_3  , nfactu_3  , fich_3, nombfich_3 , fcom_3, ftraslad_3, docref_3, tidocr_3, sfactur_3, nfactur_3, nomb_3    ,direcc_3 , chofer_3, placa_3, ruc_3, motivo_3, tienda_3, refactu_3, ordenc_3, ppartida_3, ppllegad_3, destino_3,fase_3 , subfase_3 , flag_3, almreg_3, fasereg_3, origen_3, emp_3, usuario_3, fecha_3   , hora_3  , nsalida_3, situa_3, almac_3, sfase_3, trans_3  , color_3, cerrado_3, ordens_3, habavi_3, fhanul_3, motanul_3, usuanul_3, pcanul_3, afactu_4, usrvis_3, tvisado_3, direc_3, ttrans_3, certificado_3, Brevete_3, DniChofer_3,npedido_3,glosadoc_3)
				                                                   Values (@Cia_3a, @sfactu_3a, @nfactu_3a,@fich_3, @nombfich_3,@fcom_3, @fcom_32  , 0       , ''      , ''       , ''       , @nomb_3   ,@direcc_3,@chofer_3,@placa_3,@ruc_3,@motivo_3, '02'    , '-'     , ''       ,@ppartida_3,@ppllegad_3,  @fase_3 ,@orig_3, @subfase_3, 1     , '  '    ,@orig_3   , 'EXT'   ,''    ,''        , @fcom_3, '00:00:00' , @nsalida_3, 1      , ''     , ''     ,@trans_3 , ''     , 0        , @ordens_3, ''     , Null    ,'  '      , ''       , ''      , 0       , ''      , Null     , ''      , ''      , ''          ,@Brevete_3, ''         ,''       ,@glosa)", conx)
            cmd13.Parameters.AddWithValue("@Cia_3a", Trim(Label30.Text))
            cmd13.Parameters.AddWithValue("@sfactu_3a", Trim(TextBox1.Text))
            cmd13.Parameters.AddWithValue("@nfactu_3a", Trim(TextBox2.Text))
            cmd13.Parameters.AddWithValue("@fcom_3", DateTimePicker1.Value.Date)
            cmd13.Parameters.AddWithValue("@fcom_32", DateTimePicker2.Value.Date)
            cmd13.Parameters.AddWithValue("@fich_3", Trim(TextBox8.Text))
            cmd13.Parameters.AddWithValue("@nombfich_3", Trim(TextBox9.Text))
            cmd13.Parameters.AddWithValue("@orig_3", Trim(Label23.Text))
            cmd13.Parameters.AddWithValue("@fase_3", Trim(TextBox21.Text))
            cmd13.Parameters.AddWithValue("@subfase_3", Trim(TextBox24.Text))
            cmd13.Parameters.AddWithValue("@motivo_3", Trim(ComboBox1.SelectedValue.ToString))
            cmd13.Parameters.AddWithValue("@ppartida_3", Trim(TextBox3.Text))
            cmd13.Parameters.AddWithValue("@ppllegad_3", Trim(TextBox10.Text))
            cmd13.Parameters.AddWithValue("@nomb_3", Trim(TextBox5.Text))
            cmd13.Parameters.AddWithValue("@direcc_3", Trim(TextBox13.Text))
            cmd13.Parameters.AddWithValue("@chofer_3", Trim(TextBox16.Text))
            cmd13.Parameters.AddWithValue("@placa_3", Trim(TextBox14.Text))
            cmd13.Parameters.AddWithValue("@ruc_3", Trim(TextBox11.Text))
            cmd13.Parameters.AddWithValue("@Brevete_3", Trim(TextBox15.Text))
            cmd13.Parameters.AddWithValue("@ordens_3", Trim(TextBox20.Text))
            cmd13.Parameters.AddWithValue("@nsalida_3", DateTime.Now.ToString("yyyy") & Mid(Label23.Text, 1, 2) & "2" & Trim(Label3.Text))
            cmd13.Parameters.AddWithValue("@trans_3", Trim(TextBox4.Text))
            cmd13.Parameters.AddWithValue("@glosa", Trim(TextBox6.Text))
            cmd13.ExecuteNonQuery()

            Dim cmd14 As New SqlCommand("INSERT INTO custom_vianny.dbo.Qag3000 (Empr_3 , Ano_3, talm_3  , ccom_3, ncom_3, fcom_3     , gloa_3            , glob_3,fich_3 , orig_3, nombe_3, adevol_3, transf_3, tdoc_3, grm_3, fase_3  ,flag_3, subfase_3, origen_3, situa_3, Linea_3, Lectura_3, revisado_3)
			          	                                           Values (@Cia_3a,@ano_3a,@talm_3a, '2'   ,@ncom_3a,@fcom_3    , 'GUIA DE REMISION', ''    ,@fich_3,@fase_3 , ''    , 0       , 0       , @tdoc_3, @grm_3, @orig_3, 1    , @subfase_3, 'EXT', 1, '', '', 0)", conx)
            cmd14.Parameters.AddWithValue("@Cia_3a", Trim(Label30.Text))
            cmd14.Parameters.AddWithValue("@ano_3a", Trim(Label29.Text))
            cmd14.Parameters.AddWithValue("@talm_3a", Trim(Label23.Text))
            cmd14.Parameters.AddWithValue("@ncom_3a", Trim(Label3.Text))
            cmd14.Parameters.AddWithValue("@fcom_3", DateTimePicker1.Value.Date)
            cmd14.Parameters.AddWithValue("@fich_3", Trim(TextBox8.Text))
            cmd14.Parameters.AddWithValue("@orig_3", Trim(Label23.Text))
            cmd14.Parameters.AddWithValue("@tdoc_3", "009")
            cmd14.Parameters.AddWithValue("@grm_3", TextBox1.Text & "-" & TextBox2.Text)
            cmd14.Parameters.AddWithValue("@fase_3", Trim(TextBox21.Text))
            cmd14.Parameters.AddWithValue("@subfase_3", Trim(TextBox24.Text))

            cmd14.ExecuteNonQuery()

            For i = 0 To o - 1


                Dim cmd15 As New SqlCommand("INSERT INTO custom_vianny.dbo.Fap0400 (Cia_3a , sfactu_3a  , nfactu_3a, Item_3a,linea_3a , Sinon_3a,cant_3a, unid_3a  , obser1_3a, obser2_3a, obser3_3a, obser4_3a, obser5_3a, obser6_3a, obser7_3a, obser8_3a,obser9_3a, obser10_3a, cantk_3a, pedido_3a, flag_3a, unidk_3a,ordens_3a ,parti_3a,secpartida_3a,ocorte_3a,ordtej_3a,linea2_3a,pieza_3a,ps_3a,lote_3a)
				                                                           Values (@Cia_3a, @sfactu_3a ,@nfactu_3a,@Item_3a,@linea_3a, '000'   ,@cant_3, @unid_3  ,@obser1_3a, ''       , ''       , ''       , ''       , ''       , ''       , ''       , ''      , ''        ,@cant_3a, @pedido_3a, 1      ,@unid_3a ,@ordens_3a,''      , ''          ,@ocorte_3a       ,''        ,'PT0101' ,''      ,''  ,'')", conx)
                cmd15.Parameters.AddWithValue("@Cia_3a", Trim(Label30.Text))
                cmd15.Parameters.AddWithValue("@sfactu_3a", Trim(TextBox1.Text))
                cmd15.Parameters.AddWithValue("@nfactu_3a", Trim(TextBox2.Text))
                cmd15.Parameters.AddWithValue("@Item_3a", Trim(DataGridView1.Rows(i).Cells(0).Value))
                cmd15.Parameters.AddWithValue("@linea_3a", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd15.Parameters.AddWithValue("@cant_3a", Trim(DataGridView1.Rows(i).Cells(6).Value))
                cmd15.Parameters.AddWithValue("@unid_3a", Trim(DataGridView1.Rows(i).Cells(7).Value))
                cmd15.Parameters.AddWithValue("@obser1_3a", Trim(DataGridView1.Rows(i).Cells(11).Value))
                cmd15.Parameters.AddWithValue("@pedido_3a", Trim(DataGridView1.Rows(i).Cells(1).Value))
                cmd15.Parameters.AddWithValue("@ocorte_3a", Trim(DataGridView1.Rows(i).Cells(9).Value))
                cmd15.Parameters.AddWithValue("@cant_3", Trim(DataGridView1.Rows(i).Cells(4).Value))
                cmd15.Parameters.AddWithValue("@unid_3", Trim(DataGridView1.Rows(i).Cells(5).Value))
                cmd15.Parameters.AddWithValue("@ordens_3a", Trim(TextBox20.Text))
                cmd15.ExecuteNonQuery()

                Dim cmd16 As New SqlCommand("INSERT INTO custom_vianny.dbo.Qap3000 (Empr_3a, ano_3a, talm_3a , ccom_3a, ncom_3a , item_3a, linea_3a, linea2_3a, cant_3a , unid_3a, obser1_3a, obser2_3a, obser3_3a, cantk_3a, pedido_3a, flag_3a, pieza_3a, unidk_3a, ocorte_3a, ordens_3a, Paquete_3a, OrdCos_3a)
			                                                             	Values(@Cia_3a,@ano_3a, @talm_3a, '2'    , @ncom_3a,@Item_3a,@linea_3a, 'PT0101' , @cant_3,@unid_3,@obser1_3a, ''       , ''       ,@cant_3a ,@pedido_3a, 1      , ''       ,@unid_3a,@ocorte_3a, @ordens_3a,'', '')", conx)
                cmd16.Parameters.AddWithValue("@Cia_3a", Trim(Label30.Text))
                cmd16.Parameters.AddWithValue("@ano_3a", Trim(Label29.Text))
                cmd16.Parameters.AddWithValue("@talm_3a", Trim(Label23.Text))
                cmd16.Parameters.AddWithValue("@ncom_3a", Trim(Label3.Text))
                cmd16.Parameters.AddWithValue("@Item_3a", Trim(DataGridView1.Rows(i).Cells(0).Value))
                cmd16.Parameters.AddWithValue("@linea_3a", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd16.Parameters.AddWithValue("@cant_3a", Trim(DataGridView1.Rows(i).Cells(6).Value))
                cmd16.Parameters.AddWithValue("@unid_3a", Trim(DataGridView1.Rows(i).Cells(7).Value))
                cmd16.Parameters.AddWithValue("@obser1_3a", "")
                cmd16.Parameters.AddWithValue("@pedido_3a", Trim(DataGridView1.Rows(i).Cells(1).Value))
                cmd16.Parameters.AddWithValue("@ocorte_3a", Trim(DataGridView1.Rows(i).Cells(9).Value))
                cmd16.Parameters.AddWithValue("@ordens_3a", Trim(TextBox20.Text))
                cmd16.Parameters.AddWithValue("@cant_3", Trim(DataGridView1.Rows(i).Cells(4).Value))
                cmd16.Parameters.AddWithValue("@unid_3", Trim(DataGridView1.Rows(i).Cells(5).Value))
                cmd16.ExecuteNonQuery()
            Next

            Dim hj2 As New flog
            Dim hj1 As New vlog
            hj1.gmodulo = "NSA-PRODUCCION"
            hj1.gcuser = MDIParent1.Label3.Text
            If Label34.Text = "0" Then
                hj1.gaccion = 1
            Else
                hj1.gaccion = 2
            End If

            hj1.gpc = My.Computer.Name
            hj1.gfecha = DateTimePicker3.Value
            hj1.gdetalle = Trim(Label23.Text) & Trim(TextBox1.Text) & Trim(TextBox2.Text)
            hj1.gccia = Label30.Text
            hj2.insertar_log(hj1)
            'guia_taller log
            Dim hj22 As New flog
            Dim hj12 As New vlog
            hj12.gmodulo = "GUIA-PRODUCCION"
            hj12.gcuser = MDIParent1.Label3.Text

            If Label34.Text = "0" Then
                hj12.gaccion = 1
            Else
                hj12.gaccion = 2
            End If
            hj12.gpc = My.Computer.Name
            hj12.gfecha = DateTimePicker3.Value
            hj12.gdetalle = Trim(Label23.Text) & Trim(TextBox1.Text) & Trim(TextBox2.Text)
            hj12.gccia = Label30.Text
            hj22.insertar_log(hj12)
            MsgBox("SE GUARDO LA INFORMACION CORRECTAMENTE")
            Dim respuesta As DialogResult

            respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                guia_taller.TextBox1.Text = TextBox1.Text
                guia_taller.TextBox2.Text = TextBox2.Text
                guia_taller.TextBox3.Text = Label30.Text
                guia_taller.TextBox4.Text = ""
                guia_taller.ShowDialog()
            End If
            If Label33.Text = 1 Then
                Me.Close()
            Else
                RadioButton1.Checked = True

                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox8.Text = ""
                TextBox9.Text = ""
                TextBox10.Text = ""
                TextBox11.Text = ""
                TextBox12.Text = ""
                TextBox13.Text = ""
                TextBox14.Text = ""
                TextBox15.Text = ""
                TextBox16.Text = ""
                TextBox6.Text = ""
                TextBox20.Text = ""
                DataGridView1.Rows.Clear()
                RadioButton1.Checked = True
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                TextBox5.Enabled = False
                TextBox8.Enabled = False
                TextBox9.Enabled = False
                TextBox10.Enabled = False
                TextBox11.Enabled = False
                TextBox13.Enabled = False
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
                TextBox21.Text = ""
                TextBox22.Text = ""
                Button5.Enabled = False

                DataGridView5.Rows.Clear()
                correlativo()
                Dim func As New fguiasistema
                Dim fu As New vguiacabecera
                Dim fu1 As New vguiacabecera
                fu.gserie = TextBox1.Text
                fu.gccia = Label30.Text
                TextBox2.Text = func.mostrar_guia_correlativo(fu) + 1
                Select Case TextBox2.Text.Length

                    Case "1" : TextBox2.Text = "0000000" & "" & TextBox2.Text
                    Case "2" : TextBox2.Text = "000000" & "" & TextBox2.Text
                    Case "3" : TextBox2.Text = "00000" & "" & TextBox2.Text
                    Case "4" : TextBox2.Text = "0000" & "" & TextBox2.Text
                    Case "5" : TextBox2.Text = "000" & "" & TextBox2.Text
                    Case "6" : TextBox2.Text = "00" & "" & TextBox2.Text
                    Case "7" : TextBox2.Text = "0" & "" & TextBox2.Text
                    Case "8" : TextBox2.Text = TextBox2.Text
                End Select

            End If


        End If


    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            Dim ds As New det_ns_Prodc
            ds.Label2.Text = TextBox1.Text
            ds.Label3.Text = 6
            Select Case Label23.Text
                Case "0300" : ds.Label4.Text = "1"
                Case "0400" : ds.Label4.Text = "2"
                Case "0600" : ds.Label4.Text = "3"
                Case "0500" : ds.Label4.Text = "4"
            End Select
            ds.Label5.Text = Trim(Label30.Text)
            ds.ShowDialog()
        End If
    End Sub
    Dim Rsr222, Rsr21, Rsr2333, Rsr9 As SqlDataReader
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        'Dim corre As String
        'jj.gvendedor = TextBox8.Text
        'corre = fk.buscar_correo(jj)
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='ELIMINAR_PRODUCCION '"
        Dim cmd As New SqlCommand(sql, conx)
        Rs = cmd.ExecuteReader()
        Rs.Read()
        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA GUIA DE PRODUCCION</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> ANULADO <br/>
<FONT COLOR='blue'>* AREA : </FONT>" + Trim(Label23.Text) & "-" & Trim(Label25.Text) + " <br/>
<FONT COLOR='blue'>* GUIA PRODUCCION : </FONT>" + Trim(TextBox1.Text) & "-" & Trim(TextBox2.Text) + " <br/>
<FONT COLOR='blue'>* USUARIO : </FONT>" + Trim(MDIParent1.Label3.Text) + " <br/>
<FONT COLOR='blue'>* PC : </FONT>" + Trim(My.Computer.Name) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(MDIParent1.Label7.Text) + " <br/>
<FONT COLOR='blue'>* TALLER : </FONT>" + Trim(TextBox9.Text) + "<br/>
</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add(Rs(0))
            Rs.Close()
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "GUIA PRODUCCION N°" & Trim(TextBox1.Text) & "-" & Trim(TextBox2.Text)
            .Priority = System.Net.Mail.MailPriority.Normal
        End With

        With smtp
            .EnableSsl = False
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
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim cmd16 As New SqlCommand("UPDATE custom_vianny.dbo.Fag0400 SET Flag_3 = 0 WHERE Cia_3 = @Cia_3a AND SFactu_3 = @sfactu_3a AND NFactu_3 = @nfactu_3a", conx)
        cmd16.Parameters.AddWithValue("@Cia_3a", Trim(Label30.Text))
        cmd16.Parameters.AddWithValue("@sfactu_3a", Trim(TextBox1.Text))
        cmd16.Parameters.AddWithValue("@nfactu_3a", Trim(TextBox2.Text))
        cmd16.ExecuteNonQuery()

        Dim cmd17 As New SqlCommand("UPDATE custom_vianny.dbo.Fap0400 SET Flag_3a = 0 WHERE Cia_3a = @Cia_3a AND SFactu_3a = @sfactu_3a AND NFactu_3a = @nfactu_3a", conx)
        cmd17.Parameters.AddWithValue("@Cia_3a", Trim(Label30.Text))
        cmd17.Parameters.AddWithValue("@sfactu_3a", Trim(TextBox1.Text))
        cmd17.Parameters.AddWithValue("@nfactu_3a", Trim(TextBox2.Text))
        cmd17.ExecuteNonQuery()

        Dim cmd18 As New SqlCommand("UPDATE custom_vianny.dbo.Qag3000 SET Flag_3 = 0 WHERE Empr_3 =@Cia_3a AND Ano_3 = @ano_3a AND TAlm_3 = @talm_3a AND CCom_3 = '2' AND NCom_3 = @ncom_3a", conx)
        cmd18.Parameters.AddWithValue("@Cia_3a", Trim(Label30.Text))
        cmd18.Parameters.AddWithValue("@ano_3a", Trim(Label29.Text))
        cmd18.Parameters.AddWithValue("@talm_3a", Trim(Label23.Text))
        cmd18.Parameters.AddWithValue("@ncom_3a", Trim(Label3.Text))
        cmd18.ExecuteNonQuery()

        Dim cmd19 As New SqlCommand("UPDATE custom_vianny.dbo.Qap3000 SET Flag_3a = 0 WHERE Empr_3a = @Cia_3a AND Ano_3a = @ano_3a AND TAlm_3a = @talm_3a AND CCom_3a = '2' AND NCom_3a = @ncom_3a", conx)
        cmd19.Parameters.AddWithValue("@Cia_3a", Trim(Label30.Text))
        cmd19.Parameters.AddWithValue("@ano_3a", Trim(Label29.Text))
        cmd19.Parameters.AddWithValue("@talm_3a", Trim(Label23.Text))
        cmd19.Parameters.AddWithValue("@ncom_3a", Trim(Label3.Text))
        cmd19.ExecuteNonQuery()
        Dim hj2 As New flog
        Dim hj1 As New vlog
        hj1.gmodulo = "NSA-PRODUCCION"
        hj1.gcuser = MDIParent1.Label3.Text

        hj1.gaccion = 3


        hj1.gpc = My.Computer.Name
        hj1.gfecha = DateTimePicker3.Value
        hj1.gdetalle = Trim(Label23.Text) & Trim(TextBox1.Text) & Trim(TextBox2.Text)
        hj1.gccia = Label30.Text
        hj2.insertar_log(hj1)
        MsgBox("SE ANULO CORRECTAMENTE LA GUIA")
        ENVIAR_CORREO()
        RadioButton1.Checked = True

        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox6.Text = ""
        DataGridView1.Rows.Clear()
        RadioButton1.Checked = True
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox5.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        TextBox11.Enabled = False
        TextBox13.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TextBox21.Text = ""
        TextBox22.Text = ""
        Button5.Enabled = False

        DataGridView5.Rows.Clear()
        correlativo()
        Dim func As New fguiasistema
        Dim fu As New vguiacabecera
        Dim fu1 As New vguiacabecera
        fu.gserie = TextBox1.Text
        fu.gccia = Label30.Text
        TextBox2.Text = func.mostrar_guia_correlativo(fu) + 1
        Select Case TextBox2.Text.Length

            Case "1" : TextBox2.Text = "0000000" & "" & TextBox2.Text
            Case "2" : TextBox2.Text = "000000" & "" & TextBox2.Text
            Case "3" : TextBox2.Text = "00000" & "" & TextBox2.Text
            Case "4" : TextBox2.Text = "0000" & "" & TextBox2.Text
            Case "5" : TextBox2.Text = "000" & "" & TextBox2.Text
            Case "6" : TextBox2.Text = "00" & "" & TextBox2.Text
            Case "7" : TextBox2.Text = "0" & "" & TextBox2.Text
            Case "8" : TextBox2.Text = TextBox2.Text
        End Select
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        guia_taller.TextBox1.Text = TextBox1.Text
        guia_taller.TextBox2.Text = TextBox2.Text
        guia_taller.TextBox3.Text = Label30.Text
        guia_taller.TextBox4.Text = ""
        guia_taller.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox("LA FUNCION ELIMINAR NO ESTA HABILITADA")
        'abrir()
        'Dim agregar As String = "DELETE FROM custom_vianny.dbo.Fap0400 Where CIA_3a = '" + Trim(Label30.Text) + "' And sfactu_3a = '" + Trim(TextBox1.Text) + "' And nfactu_3a = '" + Trim(TextBox2.Text) + "'"
        'Dim agregar11 As String = "DELETE FROM custom_vianny.dbo.Fag0400 Where CIA_3 = '" + Trim(Label30.Text) + "' And sfactu_3 = '" + Trim(TextBox1.Text) + "' And nfactu_3 = '" + Trim(TextBox2.Text) + "' "
        'Dim agregar1 As String = "DELETE FROM custom_vianny.dbo.Fat0400 Where Cia_3t = '" + Trim(Label30.Text) + "' And sfactu_3t = '" + Trim(TextBox1.Text) + "' And nfactu_3t = '" + Trim(TextBox2.Text) + "'"
        'Dim agregar2 As String = "DELETE FROM custom_vianny.dbo.Fah0400 Where CIA_3h = '" + Trim(Label30.Text) + "' And sfactu_3h = '" + Trim(TextBox1.Text) + "' And nfactu_3h = '" + Trim(TextBox2.Text) + "'"
        'Dim agregar3 As String = "DELETE FROM custom_vianny.dbo.Fad0400 Where CIA_3d = '" + Trim(Label30.Text) + "' And sfactu_3d = '" + Trim(TextBox1.Text) + "' And nfactu_3d = '" + Trim(TextBox2.Text) + "'"
        'Dim agregar4 As String = "DELETE FROM custom_vianny.dbo.Fat0400 Where Cia_3t = '" + Trim(Label30.Text) + "' And sfactu_3t = '" + Trim(TextBox1.Text) + "' And nfactu_3t = '" + Trim(TextBox2.Text) + "'"
        'Dim agregar5 As String = "DELETE FROM custom_vianny.dbo.Faq0400 Where CIA_4q = '" + Trim(Label30.Text) + "' And sfactu_4q = '" + Trim(TextBox1.Text) + "' And nfactu_4q = '" + Trim(TextBox2.Text) + "'"
        'Dim agregar6 As String = "DELETE FROM custom_vianny.dbo.Qap3000 Where Empr_3a = '" + Trim(Label30.Text) + "' And Ano_3a = '" + Trim(Label29.Text) + "' And talm_3a = '" + Label23.Text + "' And ccom_3a = '2' And ncom_3a = '" + Label3.Text + "'"
        'Dim agregar7 As String = "DELETE FROM custom_vianny.dbo.Qat3000 Where Empr_3t = '" + Trim(Label30.Text) + "' And Ano_3t = '" + Trim(Label29.Text) + "' And talm_3t = '" + Label23.Text + "' And ccom_3t = '2' And ncom_3t = '" + Label3.Text + "'"
        'Dim agregar8 As String = "DELETE FROM custom_vianny.dbo.QAD3000 Where Empr_3d = '" + Trim(Label30.Text) + "' And Ano_3d = '" + Trim(Label29.Text) + "' And talm_3d = '" + Label23.Text + "' And ccom_3d = '2' And ncom_3d = '" + Label3.Text + "'"
        'Dim agregar9 As String = "DELETE FROM custom_vianny.dbo.QAD3000 Where Empr_3d = '" + Trim(Label30.Text) + "' And Ano_3d = '" + Trim(Label29.Text) + "' And talm_3d = '" + Label23.Text + "' And ccom_3d = '2' And ncom_3d = '" + Label3.Text + "'"
        'Dim agregar10 As String = "DELETE FROM custom_vianny.dbo.Qag3000 Where Empr_3 = '" + Trim(Label30.Text) + "' And Ano_3 = '" + Trim(Label29.Text) + "' And talm_3 = '" + Label23.Text + "' And ccom_3 = '2' And ncom_3 = '" + Label3.Text + "'"
        'ELIMINAR(agregar)
        'ELIMINAR(agregar1)
        'ELIMINAR(agregar2)
        'ELIMINAR(agregar3)
        'ELIMINAR(agregar4)
        'ELIMINAR(agregar5)
        'ELIMINAR(agregar6)
        'ELIMINAR(agregar7)
        'ELIMINAR(agregar8)
        'ELIMINAR(agregar9)
        'ELIMINAR(agregar10)
    End Sub

    Private Sub TextBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox25.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim VAL As Integer
            Select Case Trim(TextBox25.Text).Length
                Case "1" : TextBox25.Text = "01" & "0000000" & TextBox25.Text
                Case "2" : TextBox25.Text = "01" & "000000" & TextBox25.Text
                Case "3" : TextBox25.Text = "01" & "00000" & TextBox25.Text
                Case "4" : TextBox25.Text = "01" & "0000" & TextBox25.Text
                Case "5" : TextBox25.Text = "01" & "000" & TextBox25.Text
                Case "6" : TextBox25.Text = "01" & "00" & TextBox25.Text
                Case "7" : TextBox25.Text = "01" & "0" & TextBox25.Text
                Case "8" : TextBox25.Text = "01" & TextBox25.Text
                Case "9" : TextBox25.Text = "0" & TextBox25.Text
            End Select
            'validar si op esta cerrafa
            Dim fim As Integer
            Dim sql103 As String = "select  finped_3 from custom_vianny.dbo.qag0300 where ncom_3 ='" + Trim(TextBox15.Text) + "' and  ccia ='" + Trim(Label4.Text) + "'"
            Dim cmd103 As New SqlCommand(sql103, conx)
            Rsr2333 = cmd103.ExecuteReader()
            If Rsr2333.Read() = True Then
                fim = Rsr2333(0)
            End If
            Rsr2333.Close()
            'fin de validar op cerrada


            If fim = 0 Then
                Dim opYaExiste As Boolean = False

                ' Recorrer todas las filas del DataGridView
                For Each row As DataGridViewRow In DataGridView1.Rows
                    ' Verificar si la columna que contiene la OP es igual a la opIngresada
                    If row.Cells("Op").Value IsNot Nothing AndAlso row.Cells("Op").Value.ToString() = TextBox25.Text.ToString().Trim() Then
                        opYaExiste = True
                        Exit For
                    End If
                Next

                ' Si la OP ya existe, mostrar un mensaje
                If opYaExiste Then
                    MessageBox.Show("La OP ingresada ya existe en el listado.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Else
                    Dim sql102 As String = "select ncom_3,linea_3,isnull(a.nomb_17,q.descri_3),isnull( a.unid_17,'UND')  from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + Trim(TextBox25.Text) + "' and q.ccia ='" + Trim(Label30.Text) + "'
"
                    Dim cmd102 As New SqlCommand(sql102, conx)
                    Rsr222 = cmd102.ExecuteReader()
                    Dim i5 As Integer
                    i5 = DataGridView1.Rows.Count

                    If Rsr222.Read() = True Then
                        DataGridView1.Rows.Add()
                        DataGridView1.Rows(i5).Cells(1).Value = Rsr222(0)
                        DataGridView1.Rows(i5).Cells(2).Value = Rsr222(1)
                        DataGridView1.Rows(i5).Cells(3).Value = Rsr222(2)
                        DataGridView1.Rows(i5).Cells(5).Value = Rsr222(3)
                        DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                        VAL = DataGridView1.Rows(i5).Cells(0).Value
                        Select Case VAL.ToString.Length
                            Case "1" : DataGridView1.Rows(i5).Cells(0).Value = "00" & "" & VAL
                            Case "2" : DataGridView1.Rows(i5).Cells(0).Value = "0" & "" & VAL
                            Case "3" : DataGridView1.Rows(i5).Cells(0).Value = VAL
                        End Select
                        If Label23.Text = "0400" Or Label23.Text = "0700" Then

                            abrir()
                            llenar_combo_box4(Trim(DataGridView1.Rows(i5).Cells(1).Value))
                        End If
                        TextBox25.Text = ""
                        DataGridView1.CurrentCell = DataGridView1.Rows(i5).Cells(4)
                        DataGridView1.BeginEdit(True)
                        Rsr222.Close()
                    Else
                        MsgBox("LA OP INGRESADA NO EXISTE")
                    End If
                    Rsr222.Close()
                End If
            Else
            End If
        End If
    End Sub

    Sub llenar_combo_box4(jk As String)
        Try

            Dim lsQuery As String = "select ocorte_3cg from custom_vianny.dbo.Qag300Cc where pedido_3cg ='" + jk + "'"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            pak.DataSource = loDataTable

            pak.DisplayMember = "ocorte_3cg"
            pak.ValueMember = "ocorte_3cg"
            pak.DropDownWidth = 90
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box3()
        Try

            Dim lsQuery As String = "select ocorte_3cg from custom_vianny.dbo.Qag300Cc where pedido_3cg ='0100020260'"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            pak.DataSource = loDataTable

            pak.DisplayMember = "ocorte_3cg"
            pak.ValueMember = "ocorte_3cg"
            pak.DropDownWidth = 90
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click

    End Sub

    Private Sub TextBox2_DoubleClick(sender As Object, e As EventArgs) Handles TextBox2.DoubleClick
        Dim func As New fguiasistema
        Dim fu As New vguiacabecera
        Dim fu1 As New vguiacabecera
        fu.gserie = TextBox1.Text
        fu.gccia = Label30.Text
        TextBox2.Text = func.mostrar_guia_correlativo(fu) + 1
        Select Case TextBox2.Text.Length

            Case "1" : TextBox2.Text = "0000000" & "" & TextBox2.Text
            Case "2" : TextBox2.Text = "000000" & "" & TextBox2.Text
            Case "3" : TextBox2.Text = "00000" & "" & TextBox2.Text
            Case "4" : TextBox2.Text = "0000" & "" & TextBox2.Text
            Case "5" : TextBox2.Text = "000" & "" & TextBox2.Text
            Case "6" : TextBox2.Text = "00" & "" & TextBox2.Text
            Case "7" : TextBox2.Text = "0" & "" & TextBox2.Text
            Case "8" : TextBox2.Text = TextBox2.Text
        End Select


        abrir()
        llenar_combo_box()

        TextBox2.Select()
        RadioButton1.Checked = True
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox5.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        TextBox11.Enabled = False
        TextBox13.Enabled = False
        TextBox21.Enabled = False

        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox6.Text = ""
        DataGridView1.Rows.Clear()
        TextBox21.Text = ""
        TextBox22.Text = ""
        Button5.Enabled = False
        correlativo()
        'ComboBox1.SelectedIndex = 0
        ''ComboBox2.SelectedIndex = 0
        'ComboBox3.SelectedIndex = 0
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
    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        Try
            NumConFrac(TextBox4, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        Try
            NumConFrac(TextBox4, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged

    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged

    End Sub

    Private Sub TextBox25_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox25.KeyPress
        Try
            NumConFrac(TextBox25, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox21_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox21.KeyPress
        Try
            NumConFrac(TextBox21, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        Try
            NumConFrac(TextBox2, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Try
            NumConFrac(TextBox1, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown

    End Sub

    Private Sub TextBox2_MouseUp(sender As Object, e As MouseEventArgs) Handles TextBox2.MouseUp

    End Sub

    Private Sub TextBox25_Paint(sender As Object, e As PaintEventArgs) Handles TextBox25.Paint

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        log.Label1.Text = Trim(Label23.Text) & Trim(TextBox1.Text) & Trim(TextBox2.Text)
        log.Label2.Text = "GUIA-PRODUCCION"
        log.Label3.Text = Label30.Text
        log.Show()
    End Sub
    Dim Rsr As SqlDataReader
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cantidad" Then
            Dim va, va1, vatot, final, ve As Integer
            Dim alm As String = ""
            va = 0
            va1 = 0
            If Label23.Text.ToString().Trim() <> "0300" Then

                Select Case Label23.Text.ToString().Trim()

                    Case "0700" : alm = "0300"
                    Case "0400" : alm = "0300"
                    Case "0600" : alm = "0400"
                    Case "0500"
                        Dim sql As String = "select COUNT(ncom_3a) from custom_vianny.dbo.Qap3000 where pedido_3a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()) + "' and talm_3a ='0600' and ccom_3a ='2' and flag_3a ='1'"
                        Dim cmd1 As New SqlCommand(sql, conx)
                        Rsr = cmd1.ExecuteReader()
                        If Rsr.Read() = True Then
                            ve = Convert.ToInt32(Rsr(0))
                        End If
                        Rsr.Close()
                        If ve = 0 Then
                            alm = "0400"
                        Else
                            alm = "0600"
                        End If

                End Select

                Dim sql1023 As String = "select ISNULL( SUM(cantk_3a),0) from custom_vianny.dbo.Qap3000 where pedido_3a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()) + "' and talm_3a ='" + alm + "' and Empr_3a ='" + Trim(Label30.Text) + "' and ccom_3a ='1' and flag_3a ='1'"
                Dim cmd1023 As New SqlCommand(sql1023, conx)
                Rsr233 = cmd1023.ExecuteReader()
                If Rsr233.Read() = True Then
                    va = Convert.ToInt32(Rsr233(0))
                End If
                Rsr233.Close()
                Dim sql10233 As String = "select ISNULL( SUM(cantk_3a),0) from custom_vianny.dbo.Qap3000 where pedido_3a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()) + "' and talm_3a ='" + Label23.Text.ToString().Trim() + "' and Empr_3a ='" + Trim(Label30.Text) + "' and ccom_3a ='2' and flag_3a ='1'"
                Dim cmd10233 As New SqlCommand(sql10233, conx)
                Rsr9 = cmd10233.ExecuteReader()
                If Rsr9.Read() = True Then
                    va1 = Convert.ToInt32(Rsr9(0))
                End If
                Rsr9.Close()
                'MsgBox(va.ToString())
                'MsgBox(va1.ToString())
                final = va - va1
                vatot = Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
                If vatot > final Then
                    MsgBox("LA CANTIDAD INGRESADA ES MAYOR A LOS INGRESADO EN EL PROCESO ANTERIOR --> EN EL PROCESO ANTERIO INGRESO  " + Convert.ToString(final) + " PRENDAS PARA REGULARIZAR LA OP ")
                    DataGridView1.Rows(e.RowIndex).Cells(6).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim func As New fguiasistema
            Dim fu As New vguiacabecera
            Dim fu1 As New vguiacabecera
            fu.gserie = TextBox1.Text
            fu.gccia = Label30.Text
            TextBox2.Text = func.mostrar_guia_correlativo(fu) + 1
            Select Case TextBox2.Text.Length

                Case "1" : TextBox2.Text = "0000000" & "" & TextBox2.Text
                Case "2" : TextBox2.Text = "000000" & "" & TextBox2.Text
                Case "3" : TextBox2.Text = "00000" & "" & TextBox2.Text
                Case "4" : TextBox2.Text = "0000" & "" & TextBox2.Text
                Case "5" : TextBox2.Text = "000" & "" & TextBox2.Text
                Case "6" : TextBox2.Text = "00" & "" & TextBox2.Text
                Case "7" : TextBox2.Text = "0" & "" & TextBox2.Text
                Case "8" : TextBox2.Text = TextBox2.Text
            End Select

            'fu1.gmes = DateTime.Now.ToString("MM")
            'fu1.galmacen = Label23.Text
            'fu1.gano = Label29.Text
            'fu1.gccia = Label30.Text
            'Label3.Text = func.mostrar_correlativo_nsa(fu1) + 1
            'Select Case Label3.Text.Length

            '    Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
            '    Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
            '    Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
            '    Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
            'End Select
            abrir()
            llenar_combo_box()
            'llenar_combo_box2(ComboBox2, Label23.Text)
            TextBox2.Select()
            RadioButton1.Checked = True
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            TextBox5.Enabled = False
            TextBox8.Enabled = False
            TextBox9.Enabled = False
            TextBox10.Enabled = False
            TextBox11.Enabled = False
            TextBox13.Enabled = False
            TextBox21.Enabled = False

            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
            TextBox11.Text = ""
            TextBox12.Text = ""
            TextBox13.Text = ""
            TextBox14.Text = ""
            TextBox15.Text = ""
            TextBox16.Text = ""
            TextBox6.Text = ""
            DataGridView1.Rows.Clear()
            TextBox21.Text = ""
            TextBox22.Text = ""
            Button5.Enabled = False
            correlativo()
            'ComboBox1.SelectedIndex = 0
            ''ComboBox2.SelectedIndex = 0
            'ComboBox3.SelectedIndex = 0
        End If

    End Sub

    Private Sub TextBox21_Validating(sender As Object, e As CancelEventArgs) Handles TextBox21.Validating
        If Trim(DirectCast(sender, TextBox).Text).Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL DESTINO")
        End If
    End Sub

    Private Sub TextBox8_Validating(sender As Object, e As CancelEventArgs) Handles TextBox8.Validating
        If Trim(DirectCast(sender, TextBox).Text).Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL RUC")
        End If
    End Sub

    Private Sub TextBox4_Validating(sender As Object, e As CancelEventArgs) Handles TextBox4.Validating
        If Trim(DirectCast(sender, TextBox).Text).Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL TRANSPORTE")
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Obs" Then
            OBSERVACION.Label2.Text = 3
            OBSERVACION.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value)
            OBSERVACION.Label3.Text = e.RowIndex
            OBSERVACION.ShowDialog()
        End If
    End Sub

    Private Sub TextBox8_MouseDown(sender As Object, e As MouseEventArgs) Handles TextBox8.MouseDown

    End Sub
End Class