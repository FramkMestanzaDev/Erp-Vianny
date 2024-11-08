Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Net.Mail
Public Class Guia_remision
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
    Public comando As SqlCommand
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
    Sub llenar_combo_box(ByVal cb As ComboBox)
        Try

            conn = New SqlDataAdapter("select rtrim(ltrim(motiv_17f)) AS motiv_17f,rtrim(ltrim(nomb_17f)) AS nomb_17f from custom_vianny.dbo.Fag1700 where ccia_17F ='" + Trim(Label30.Text) + "'", conx)
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
    Sub llenar_combo_box2(ByVal cb As ComboBox, ByVal alm As String)
        Try

            conn = New SqlDataAdapter("select dest_21s,rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where  Empr_21s =" + Trim(Label30.Text) + "AND almac_21s=" + alm, conx)
            conn.Fill(ty2)
            ComboBox2.DataSource = ty2
            ComboBox2.DisplayMember = "nomb_21s"
            ComboBox2.ValueMember = "dest_21s"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ''Orden_Packing.DataGridView1.Columns(6).Visible = True
        'Orden_Packing.Show()
        If Trim(TextBox8.Text).Length = 0 Then
            MsgBox("PRIMERO DE INGRESAR EL CLIENTE")
        Else
            Pedidos.Label3.Text = Label29.Text
            Pedidos.Label4.Text = 1
            Pedidos.Label5.Text = Label30.Text
            Pedidos.Label6.Text = TextBox14.Text
            Pedidos.Label7.Text = TextBox6.Text
            Pedidos.Show()
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
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        Dim camposFaltantes As New List(Of String) ' Lista para almacenar los nombres de los campos faltantes
            ' Validar campos obligatorios
            If String.IsNullOrEmpty(TextBox8.Text) Then
                camposFaltantes.Add(" Cliente ")
            End If

        If camposFaltantes.Count > 0 Then
            ' Mostrar mensaje de error indicando los campos faltantes
            Dim camposFaltantesStr As String = String.Join(", ", camposFaltantes)
            MessageBox.Show("Falta ingresar información en los siguientes campos: " & camposFaltantesStr, "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim articulos As New Articulos_stock
            articulos.Label3.Text = Label23.Text
            articulos.Label4.Text = 1
            articulos.Label5.Text = Label29.Text
            articulos.Label6.Text = Label30.Text
            articulos.Label7.Text = 2
            articulos.Label8.Text = Trim(TextBox1.Text) & Trim(TextBox2.Text)
            articulos.ShowDialog()
        End If



    End Sub
    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        ToolTip1.SetToolTip(PictureBox1, "IMPRIMIR INFORMACION")
        ToolTip1.ToolTipTitle = "GUIA REMISION"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub
    Private Sub PictureBox2_MouseHover(sender As Object, e As EventArgs) Handles PictureBox2.MouseHover
        ToolTip2.SetToolTip(PictureBox2, "EDITAR INFORMACION")
        ToolTip2.ToolTipTitle = "GUIA REMISION"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub
    Private Sub PictureBox3_MouseHover(sender As Object, e As EventArgs)
        ToolTip3.SetToolTip(Button5, "GUARDAR INFORMACION")
        ToolTip3.ToolTipTitle = "GUIA REMISION"
        ToolTip3.ToolTipIcon = ToolTipIcon.Info
    End Sub
    Private Sub PictureBox4_MouseHover(sender As Object, e As EventArgs) Handles PictureBox4.MouseHover
        ToolTip4.SetToolTip(PictureBox4, "AGREGAR INFORMACION")
        ToolTip4.ToolTipTitle = "GUIA REMISION"
        ToolTip4.ToolTipIcon = ToolTipIcon.Info
    End Sub
    Private Sub PictureBox5_MouseHover(sender As Object, e As EventArgs) Handles PictureBox5.MouseHover
        ToolTip5.SetToolTip(PictureBox5, "ANULAR INFORMACION")
        ToolTip5.ToolTipTitle = "GUIA REMISION"
        ToolTip5.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub Guia_remision_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        fu1.gmes = DateTime.Now.ToString("MM")
        fu1.galmacen = Label23.Text
        fu1.gano = Label29.Text
        fu1.gccia = Label30.Text
        Label3.Text = func.mostrar_correlativo_nsa(fu1) + 1
        Select Case Label3.Text.Length

            Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
            Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
            Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
            Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
        End Select
        abrir()
        llenar_combo_box(ComboBox1)
        llenar_combo_box2(ComboBox2, Label23.Text)
        TextBox2.Select()

        RadioButton1.Checked = True
        RadioButton4.Checked = True
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
        ComboBox3.Enabled = False
        PictureBox1.Enabled = False
        PictureBox2.Enabled = False
        Button5.Enabled = False
        PictureBox4.Enabled = False
        PictureBox5.Enabled = False
        ComboBox1.SelectedIndex = 0
        'ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 1
    End Sub
    Dim dt1, dt2 As New DataTable

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

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        sieditable()
        TextBox3.Enabled = True
        DataGridView1.Enabled = True
        Button1.Enabled = True
        TextBox20.Enabled = True
    End Sub
    Dim Rsr3001, Rsr22, Rsr30014 As SqlDataReader
    Public Sub sieditable()
        Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label30.Text + "' AND A.talm_3 ='" + Label23.Text + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr3001 = cmd102213.ExecuteReader()
        Dim jo As Integer
        If Rsr3001.Read() Then
            jo = Rsr3001(0)
        Else
            jo = 0
        End If
        Rsr3001.Close()
        Dim sql1022134 As String = "select COUNT( nfactu_3) from custom_vianny.dbo.fag0400 where  CIA_3 =" + Label30.Text + " and sfactu_3 ='" + Trim(TextBox1.Text) + "' and nfactu_3 =" + Trim(TextBox2.Text) + " and cerrado_3 ='2'"
        Dim cmd1022134 As New SqlCommand(sql1022134, conx)
        Rsr30014 = cmd1022134.ExecuteReader()
        Dim jo4 As Integer
        If Rsr30014.Read() Then
            jo4 = Rsr30014(0)
        Else
            jo4 = 0
        End If
        Rsr30014.Close()
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("QUIERES EDITAR ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            If jo = 0 Then

                If jo4 = 0 Then
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
                    PictureBox1.Enabled = True
                    Button5.Enabled = True
                    PictureBox4.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox3.Enabled = True
                    Label32.Text = "1"
                Else
                    MsgBox("LA GUIA DE REMISION YA ESTA ENVIADA A SUNAT - NO SE PEUDE MODIFICAR ")

                End If

            Else
                MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
            End If
        End If

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    'Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    '    MsgBox(ComboBox1.SelectedValue.ToString)
    'End Sub
    Dim hg As DataTable
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        ty2.Clear()
        ty.Clear()
        abrir()
        llenar_combo_box(ComboBox1)
        llenar_combo_box2(ComboBox2, Label23.Text)
        Dim i As Integer
        Dim ml As New vguiacabecera
        Dim ml1 As New vguiacabecera
        Dim mk As New fguiasistema
        i = Len(TextBox2.Text)
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        If e.KeyCode = Keys.Enter Then
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
                If mk.mostrar_almacen_guia(ml) = Label23.Text Then
                    Dim jk As New fguiasistema

                    dt1 = jk.mostrar_cabecera_guia(ml)
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
                        If Trim(DataGridView2.Rows(0).Cells(25).Value) = "003" Then
                            TextBox18.Text = "BOL"
                        Else
                            TextBox18.Text = "FAC"
                        End If

                        TextBox7.Text = DataGridView2.Rows(0).Cells(6).Value
                        TextBox17.Text = DataGridView2.Rows(0).Cells(7).Value
                        Label3.Text = DataGridView2.Rows(0).Cells(11).Value & "" & DataGridView2.Rows(0).Cells(12).Value
                        DateTimePicker1.Value = DataGridView2.Rows(0).Cells(20).Value
                        DateTimePicker2.Value = DataGridView2.Rows(0).Cells(21).Value
                        TextBox24.Text = DataGridView2.Rows(0).Cells(22).Value
                        TextBox23.Text = DataGridView2.Rows(0).Cells(23).Value
                        TextBox20.Text = Trim(DataGridView2.Rows(0).Cells(19).Value)
                        If Trim(DataGridView2.Rows(0).Cells(24).Value) = "1" Then
                            RadioButton4.Checked = True
                            RadioButton3.Checked = False
                        Else
                            RadioButton4.Checked = False
                            RadioButton3.Checked = True
                        End If
                        abrir()
                        enunciado = New SqlCommand("select nomb_18 as EMPRESA from custom_vianny.dbo.Fag1800 where empr_18 =" + Trim(Label30.Text) + " and  trans_18 =" + TextBox4.Text, conx)
                        respuesta = enunciado.ExecuteReader
                        While respuesta.Read
                            TextBox5.Text = respuesta.Item("EMPRESA")
                        End While
                        respuesta.Close()
                        enunciado1 = New SqlCommand("select rtrim(ltrim(nomb_17f)) AS nomb_17f from custom_vianny.dbo.Fag1700 where ccia_17F ='" + Trim(Label30.Text) + "' and motiv_17f = '" + DataGridView2.Rows(0).Cells(15).Value + "'", conx)
                        respuesta1 = enunciado1.ExecuteReader
                        While respuesta1.Read
                            ComboBox1.Text = respuesta1.Item("nomb_17f")
                        End While
                        respuesta1.Close()
                        enunciado2 = New SqlCommand("select rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where   Empr_21s ='" + Trim(Label30.Text) + "' AND almac_21s='" + Label23.Text + "' and dest_21s='" + DataGridView2.Rows(0).Cells(16).Value + "'", conx)
                        respuesta2 = enunciado2.ExecuteReader
                        While respuesta2.Read
                            ComboBox2.Text = respuesta2.Item("nomb_21s")
                        End While
                        respuesta2.Close()
                        If DataGridView2.Rows(0).Cells(15).Value = "EXT" Then
                            ComboBox3.Text = "EXTERNO"
                        Else
                            If DataGridView2.Rows(0).Cells(15).Value = "INT" Then
                                ComboBox3.Text = "MISMA EMPRESA"
                            End If
                        End If
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
                            DataGridView1.Rows(p).Cells(6).Value = DataGridView3.Rows(p).Cells(7).Value
                            DataGridView1.Rows(p).Cells(7).Value = DataGridView3.Rows(p).Cells(1).Value
                            DataGridView1.Rows(p).Cells(8).Value = DataGridView3.Rows(p).Cells(2).Value
                            DataGridView1.Rows(p).Cells(10).Value = DataGridView3.Rows(p).Cells(7).Value
                            DataGridView1.Rows(p).Cells(17).Value = Trim(DataGridView3.Rows(p).Cells(10).Value)
                            DataGridView1.Rows(p).Cells(18).Value = Trim(DataGridView3.Rows(p).Cells(11).Value)
                            DataGridView1.Rows(p).Cells(19).Value = Trim(DataGridView3.Rows(p).Cells(12).Value)
                            DataGridView1.Rows(p).Cells(20).Value = Trim(DataGridView3.Rows(p).Cells(13).Value)
                            DataGridView1.Rows(p).Cells(21).Value = Trim(DataGridView3.Rows(p).Cells(14).Value)
                            DataGridView1.Rows(p).Cells(9).Value = Trim(DataGridView3.Rows(p).Cells(15).Value)
                            Select Case Trim(DataGridView3.Rows(p).Cells(16).Value)
                                Case "1" : DataGridView1.Rows(p).Cells(22).Value = "VERDE"
                                Case "2" : DataGridView1.Rows(p).Cells(22).Value = "AMARILLO"
                                Case "3" : DataGridView1.Rows(p).Cells(22).Value = "CELESTE"
                                Case "4" : DataGridView1.Rows(p).Cells(22).Value = "ROJO"
                                Case "" : DataGridView1.Rows(p).Cells(22).Value = ""
                            End Select
                            Dim fh As New fguiasistema
                            Dim gb As New vguiacabecera
                            gb.gccia = Label30.Text
                            gb.gCodArtIni = DataGridView3.Rows(p).Cells(5).Value
                            gb.galmacen = Label23.Text
                            If DataGridView1.Rows(p).Cells(6).Value = "" Then
                                gb.gFiltroDescrip = "null"
                                gb.gModal = 1
                            Else
                                gb.gFiltroDescrip = DataGridView1.Rows(p).Cells(6).Value
                                gb.gModal = 2
                            End If
                            gb.gano = Label29.Text
                            hg = fh.stock_guia(gb)
                            If hg.Rows.Count > 0 Then
                                DataGridView4.DataSource = hg
                                DataGridView1.Rows(p).Cells(11).Value = DataGridView4.Rows(0).Cells(10).Value + DataGridView1.Rows(p).Cells(7).Value
                            Else
                                DataGridView1.Rows(p).Cells(11).Value = 0.00 + DataGridView1.Rows(p).Cells(7).Value
                            End If

                            'MsgBox(DataGridView4.Rows(0).Cells(10).Value)

                        Next
                        DataGridView1.Columns(0).Frozen = True
                        DataGridView1.Columns(1).Frozen = True
                        DataGridView1.Columns(2).Frozen = True

                    End If
                    'noeditable()
                    DataGridView1.Enabled = False

                    PictureBox2.Enabled = True
                    PictureBox5.Enabled = True
                    Button1.Enabled = False
                    If RadioButton2.Checked = True Then
                        Label15.Visible = True
                    Else
                        Label15.Visible = False
                    End If

                    'RadioButton2.Checked = True
                    'RadioButton1.Checked = True
                    PictureBox1.Enabled = True
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
                Dim func As New fguiasistema

                Dim fu1 As New vguiacabecera
                fu1.gmes = DateTime.Now.ToString("MM")
                fu1.galmacen = Label23.Text
                fu1.gano = Label29.Text
                fu1.gccia = Label30.Text
                Label3.Text = func.mostrar_correlativo_nsa(fu1) + 1
                Select Case Label3.Text.Length

                    Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
                    Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
                    Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
                    Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
                End Select
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                TextBox8.Enabled = True
                TextBox10.Enabled = True
                TextBox11.Enabled = True
                TextBox13.Enabled = True
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                ComboBox3.Enabled = True
                PictureBox1.Enabled = True

                Button5.Enabled = True
                PictureBox4.Enabled = True


                TextBox4.Text = ""
                TextBox18.Text = ""
                TextBox7.Text = ""
                TextBox17.Text = ""

                ComboBox1.Text = ""
                ComboBox2.Text = ""
                ComboBox3.Text = ""
                If Label35.Text = "1" Then
                    mostrar_automaticoguia()
                    Mostar_stock()
                End If
            End If
            TextBox4.Select()
        Else
            If e.KeyCode = Keys.F1 Then
                Dim det As New Detalle_guia
                det.TextBox2.Text = TextBox1.Text
                det.Label3.Text = Label23.Text
                det.Label3.Text = Label23.Text
                det.ShowDialog()
            End If

        End If
    End Sub
    Dim Rsr2, Rsr25 As SqlDataReader
    Sub mostrar_automaticoguia()

        Dim sql102 As String = "SELECT OpDtm,CopDtm,c.nomb_17,CanDtm,c.unid_17,CodDtm,LoPDtm,CalDtm FROM DetalleTelaManufactura d left join custom_vianny.dbo.cag1700  c on c.ccia ='01' and d.CopDtm = c.linea_17 where EstDtm ='1'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        Dim i5 As Integer
        i5 = 0

        While Rsr2.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i5).Cells(1).Value = Rsr2(0)
            DataGridView1.Rows(i5).Cells(2).Value = Rsr2(1)
            DataGridView1.Rows(i5).Cells(3).Value = Rsr2(2)
            DataGridView1.Rows(i5).Cells(7).Value = Rsr2(3)
            DataGridView1.Rows(i5).Cells(8).Value = Rsr2(4)
            DataGridView1.Rows(i5).Cells(23).Value = Rsr2(5)
            DataGridView1.Rows(i5).Cells(6).Value = Rsr2(6).ToString().Trim()
            DataGridView1.Rows(i5).Cells(10).Value = Rsr2(6).ToString().Trim()
            If Rsr2(7).ToString().Trim() = "VERDE" Then
                DataGridView1.Rows(i5).Cells(22).Value = clasif.Items(0)
            Else
                If Rsr2(7).ToString().Trim() = "AMARILLO" Then
                    DataGridView1.Rows(i5).Cells(22).Value = clasif.Items(1)
                Else
                    If Rsr2(7).ToString().Trim() = "CELESTE" Then
                        DataGridView1.Rows(i5).Cells(22).Value = clasif.Items(2)
                    Else
                        DataGridView1.Rows(i5).Cells(22).Value = clasif.Items(3)
                    End If
                End If
            End If

            DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
            i5 = i5 + 1
        End While
        Rsr2.Close()



    End Sub
    Sub Mostar_stock()
        'stock fisico

        For i8 = 0 To DataGridView1.Rows.Count - 1
            Dim sql1021 As String = "EXEC ReporteStockFisicoUnitario '" + Trim(Label30.Text) + "','" + Trim(Label29.Text) + "','" + Trim(Label23.Text) + "','" + Trim(Label29.Text) + "0101','" + Trim(Label29.Text) + "1231','" + DataGridView1.Rows(i8).Cells(2).Value + "','" + DataGridView1.Rows(i8).Cells(2).Value + "',NULL, 2,'" + DataGridView1.Rows(i8).Cells(6).Value + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr25 = cmd1021.ExecuteReader()
            If Rsr25.Read() Then
                DataGridView1.Rows(i8).Cells(11).Value = Rsr25(10)
            End If
            Rsr25.Close()
        Next

        'fin stock
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            If Trim(TextBox4.Text).Length = 0 Then
                MessageBox.Show("FALTA QUE AGREGAR EL TRANSPORTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim cli As New Clientes
                cli.TextBox3.Text = "20"
                cli.ShowDialog()
            End If

        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 12 Then
            Dim num1, num2 As Integer
            num1 = e.RowIndex
            num2 = e.ColumnIndex
            Orden_Packing.TextBox1.Text = DataGridView1.Rows(num1).Cells(6).Value

            Orden_Packing.TextBox2.Text = DataGridView1.Rows(num1).Cells(3).Value
            Orden_Packing.Label3.Text = Label23.Text
            Orden_Packing.Label4.Text = Mid(DataGridView1.Rows(num1).Cells(6).Value, 1, 6)
            Orden_Packing.TextBox4.Text = DataGridView1.Rows(num1).Cells(7).Value
            Orden_Packing.Show()
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.Text = "MISMA EMPRESA" Then
            TextBox8.Text = "20508740361"
            TextBox9.Text = "CONSORCIO TEXTIL VIANNY"
            TextBox10.Text = "JR. LOS OLMOS 358 CANTO BELLO SJL"
        End If
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
    '        .Body = "Nota Genero la Guia N°" & TextBox1.Text & "-" & TextBox2.Text
    '        .Subject = "Cliente" & TextBox9.Text & "Almacen" & Label23.Text
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
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Reporte_Guia.TextBox1.Text = TextBox1.Text
        Reporte_Guia.TextBox2.Text = TextBox2.Text
        Reporte_Guia.TextBox3.Text = Label30.Text
        Reporte_Guia.Show()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = "21"
            Clientes.Show()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox10.Text = "Av. Las Torres Mz. F Lote 8 – Huachipa - Chosica -Lima"
        Else
            If CheckBox2.Checked = False Then
                TextBox10.Text = "JR. LOS OLMOS NRO. 358 URB. CANTO BELLO (ALT. PARDERO 1 AV.CONTO GRANDE) LIMA-LIMA-SJL "
            End If
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Dim Rsr30012, Rsr212, Rsr3001296 As SqlDataReader
    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Dim respuesta As DialogResult
        Dim ml As New fguiasistema
        Dim mk As New vguiacabecera
        Dim ll As Integer
        Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label30.Text + "' AND A.talm_3 ='" + Label23.Text + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr30012 = cmd102213.ExecuteReader()
        Dim jo As Integer
        If Rsr30012.Read() Then
            jo = Rsr30012(0)
        Else
            jo = 0
        End If
        Rsr30012.Close()
        respuesta = MessageBox.Show("QUIERES ANULAR LA GUIA REMISION?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            ll = DataGridView1.Rows.Count
            If jo = 0 Then
                For i = 0 To ll - 1
                    mk.gsfactu = TextBox1.Text
                    mk.gnfactu = TextBox2.Text
                    mk.gano = Year(DateTimePicker1.Value)
                    mk.galmacen = Label23.Text
                    mk.gcodigo = DataGridView1.Rows(i).Cells(2).Value
                    mk.gpartida = DataGridView1.Rows(i).Cells(6).Value
                    mk.gcantidad = DataGridView1.Rows(i).Cells(7).Value
                    mk.gccia = Label30.Text
                    ml.anular_guia(mk)
                Next
            Else
                MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")

            End If



            MsgBox("SE ANULO LA GUIA SOLICITADA")
            Me.Close()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.F1 Then
            If Trim(TextBox2.Text).Length = 0 Then
                MessageBox.Show("EL CORRELATIVO DE LA GUIA ESTA VACIO", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim trans As New Transportistas
                trans.TextBox1.Text = 1
                trans.ShowDialog()
            End If

        Else
            If e.KeyCode = Keys.Enter Then
                abrir()
                enunciado3 = New SqlCommand("select nomb_18 , direcc_18,ruc_18,chofer_18,placa_18 + marca_18 AS PLA,brevete_18 from custom_vianny.dbo.Fag1800 where empr_18 =" + Trim(Label30.Text) + " and trans_18 =" + TextBox4.Text, conx)
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
    Dim Rsr35, Rsr215, Rsr216 As SqlDataReader
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, fila, suma As Integer

        suma = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cant. Consumo" Then
            abrir()
            Dim agregar As String = "delete  Almacen_Ficticio  WHERE  CodAfi = '" + Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) + "' AND AlmAfi ='" + Trim(Label23.Text) + "' AND EmpAfi = '" + Trim(Label30.Text) + "' AND GuiAfi ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' AND PerAfi='" + Trim(Label29.Text) + "'"
            ELIMINAR(agregar)

            For i10 = 0 To suma - 1

                If (Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString()) = Trim(DataGridView1.Rows(i10).Cells(2).Value.ToString())) Then


                    If (DataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString().Trim() = DataGridView1.Rows(i10).Cells(6).Value.ToString().Trim()) Then

                        Dim stock As Double
                        stock = 0
                        Dim sql1021 As String = "EXEC ReporteStockFisicoUnitario '" + Trim(Label30.Text) + "','" + Trim(Label29.Text) + "','" + Trim(Label23.Text) + "','" + Trim(Label29.Text) + "0101','" + Trim(Label29.Text) + "1231','" + DataGridView1.Rows(i10).Cells(2).Value + "','" + DataGridView1.Rows(i10).Cells(2).Value + "',NULL, 2,'" + DataGridView1.Rows(i10).Cells(6).Value + "'"
                        Dim cmd1021 As New SqlCommand(sql1021, conx)
                        Rsr212 = cmd1021.ExecuteReader()

                        If Rsr212.Read() Then
                            stock = Convert.ToDouble(Rsr212(10))
                        End If
                        Rsr212.Close()
                        'MsgBox(stock.ToString)
                        Dim sql1022139 As String = "select SUM(CanAfi) from Almacen_Ficticio where CodAfi='" + Trim(DataGridView1.Rows(i10).Cells(2).Value) + "' and AlmAfi='" + Trim(Label23.Text) + "' and ParAfi ='" + Trim(DataGridView1.Rows(i10).Cells(6).Value) + "'  group by CodAfi,AlmAfi,EmpAfi,ParAfi"
                        Dim cmd1022139 As New SqlCommand(sql1022139, conx)
                        Rsr3001296 = cmd1022139.ExecuteReader()
                        Dim jo As Double
                        If Rsr3001296.Read() Then
                            jo = Convert.ToDouble(Rsr3001296(0))
                        Else
                            jo = 0
                        End If
                        Rsr3001296.Close()

                        'MsgBox(jo.ToString)
                        'Registo de informacion
                        If Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(7).Value) > (stock - jo) Then
                            MsgBox("LA CANTIDAD SOLICITADA ES MAYOR AL STOCK")
                            DataGridView1.Rows(fila).Cells(7).Value = 0
                        Else
                            Dim cmd168 As New SqlCommand("insert into Almacen_Ficticio (ItmAfi,CodAfi,AlmAfi,EmpAfi,GuiAfi,PerAfi,CanAfi,ParAfi) values (@ItmAfi,@CodAfi,@AlmAfi,@EmpAfi,@GuiAfi,@PerAfi,@CanAfi,@ParAfi)", conx)
                            cmd168.Parameters.AddWithValue("@ItmAfi", Trim(DataGridView1.Rows(i10).Cells(0).Value))
                            cmd168.Parameters.AddWithValue("@CodAfi", Trim(DataGridView1.Rows(i10).Cells(2).Value))
                            cmd168.Parameters.AddWithValue("@AlmAfi", Trim(Label23.Text))
                            cmd168.Parameters.AddWithValue("@EmpAfi", Trim(Label30.Text))
                            cmd168.Parameters.AddWithValue("@GuiAfi", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                            cmd168.Parameters.AddWithValue("@PerAfi", Trim(Label29.Text))
                            cmd168.Parameters.AddWithValue("@CanAfi", Convert.ToDouble(DataGridView1.Rows(i10).Cells(7).Value))
                            cmd168.Parameters.AddWithValue("@ParAfi", Trim(DataGridView1.Rows(i10).Cells(6).Value))
                            cmd168.ExecuteNonQuery()
                        End If


                        'fin del registro




                        'If Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(7).Value) <> Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(11).Value) Then
                        '    DataGridView1.Rows(i10).Cells(11).Value = (stock - jo).ToString()
                        '    MsgBox("paso")
                        'End If



                    End If
                End If
            Next







        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "O. Tejeduria" Then
            abrir()
            Dim sql1021 As String = "select distinct ordtej_3r, pedido_3r from custom_vianny.dbo.marg0001 where partida_3r ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr22 = cmd1021.ExecuteReader()
            If Rsr22.Read() = True Then
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = Rsr22(1)
                DataGridView1.Rows(e.RowIndex).Cells(9).Value = Rsr22(0)
            Else
                DataGridView1.Rows(e.RowIndex).Cells(9).Value = ""
                DataGridView1.Rows(e.RowIndex).Cells(9).Value = ""
            End If
            Rsr22.Close()
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OP" Then
            Dim sql10221348 As String = "SELECT count(ncom_3) as CANTIDAD FROM custom_vianny.dbo.qag0300 where ncom_3 ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) + "' and ccia ='" + Label30.Text + "'"
            Dim cmd10221348 As New SqlCommand(sql10221348, conx)
            Rsr35 = cmd10221348.ExecuteReader()
            If Rsr35.Read = True Then
                If Rsr35(0) <= 0 Then
                    MessageBox.Show("lA OP INGRESADA NO EXISTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    DataGridView1.Rows(e.RowIndex).Cells(1).Value = ""
                End If

            Else
                MessageBox.Show("lA OP INGRESADA NO EXISTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = ""
            End If
            Rsr35.Close()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim ui, ct As Integer
        ct = 0
        ui = DataGridView1.Rows.Count
        For i11 = 0 To ui - 1
            If DataGridView1.Rows(i11).Cells(22).Value Is Nothing Then
                ct = ct + 1
            End If
        Next
        Dim cant, Cop, Cup, Ccon, Cpartida, Ccolor As Integer

        cant = DataGridView1.Rows.Count
        Cop = 0
        Cup = 0
        Ccon = 0
        Cpartida = 0
        Ccolor = 0
        For i1 = 0 To cant - 1
            If Trim(DataGridView1.Rows(i1).Cells(1).Value).Length = 0 Then
                Cop = Cop + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(4).Value).Length = 0 Then
                Cup = Cup + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(7).Value).Length = 0 Then
                Ccon = Ccon + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(6).Value).Length = 0 Then
                Cpartida = Cpartida + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(22).Value).Length = 0 Then
                Ccolor = Ccolor + 1
            End If
        Next
        Dim VacioCliente As Boolean
        If String.IsNullOrEmpty(TextBox8.Text) Then
            VacioCliente = True
        End If
        If Cop > 0 And Trim(ComboBox1.Text) <> "VENTA" Then
            MessageBox.Show("FALTA INGRESAR LA OP EN UNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)


        Else
            If Cup > 0 Then
                MessageBox.Show("FALTA INGRESAR LA CANTIDAD DE PRESENTACION EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Else
                If Ccon > 0 Then
                    MessageBox.Show("FALTA INGRESAR LA CANTIDAD EN KG/MTS EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Else
                    If Cpartida > 0 Then
                        MessageBox.Show("FALTA INGRESAR LA PARTIDA EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

                    Else
                        If Ccolor > 0 Then
                            MessageBox.Show("FALTA INGRESAR LA CALIDAD EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

                        Else
                            If VacioCliente = True Then
                                MessageBox.Show("FALTA INGRESAR EL CLIENTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                '--------------------NOTA SALIDA
                                Dim ns1 As New vnsacabece
                                Dim ns2 As New fnsa

                                ns1.gccia = Trim(Label30.Text)
                                ns1.gncom = Trim(Label3.Text)
                                ns1.gfecha = DateTimePicker1.Value
                                ns1.gglosa = Trim(TextBox6.Text)
                                ns1.gdoc = "009"
                                ns1.gguia = TextBox1.Text & "-" & TextBox2.Text
                                ns1.gsfactu = Trim(TextBox7.Text)
                                ns1.gnfactu = Trim(TextBox17.Text)
                                ns1.gruc = Trim(TextBox8.Text)
                                ns1.galmacen = Label23.Text
                                If ComboBox3.Text = "INTERNA" Then
                                    ns1.gtipointexte = "INT"
                                Else
                                    ns1.gtipointexte = "EXT"
                                End If
                                ns1.gorige_sali = ComboBox2.SelectedValue.ToString
                                ns1.gtdocento = "009"
                                ns1.gusuario = Trim(TextBox19.Text)
                                ns1.gano = Trim(Label29.Text)
                                ns1.gfase = Trim(TextBox24.Text)
                                ns1.gpedorig_3 = ""
                                If RadioButton4.Checked = True Then
                                    ns1.gadevol = "0"
                                Else
                                    If RadioButton3.Checked = True Then
                                        ns1.gadevol = "1"
                                    End If

                                End If

                                '--------------------------------------- GUIA REMISION
                                Dim gr1 As New vguiacabecera
                                Dim gr2 As New fguiasistema
                                Dim gr7 As New vguiacabecera
                                gr7.gsfactu = Trim(TextBox1.Text)
                                gr7.gnfactu = Trim(TextBox2.Text)
                                gr7.gncom = Trim(Label3.Text)
                                gr7.galmacen = Trim(Label23.Text)
                                gr7.gano = Trim(Label29.Text)
                                gr7.gccia = Trim(Label30.Text)
                                gr2.eliminar_guia(gr7)
                                '-----------------------------------------
                                gr1.gsfactu = Trim(TextBox1.Text)
                                gr1.gnfactu = Trim(TextBox2.Text)
                                gr1.gruc = Trim(TextBox8.Text)
                                gr1.gruc3 = Trim(TextBox11.Text)
                                gr1.gtip_documento = "001"
                                gr1.gruc_transpo = Trim(TextBox12.Text)
                                gr1.gdirec_transport = Trim(TextBox13.Text)
                                gr1.grsocial = Trim(TextBox9.Text)
                                gr1.gfcom = DateTimePicker1.Value
                                gr1.gfcom1 = DateTimePicker2.Value
                                gr1.gdireccion = Trim(TextBox10.Text)
                                gr1.gplaca = Trim(TextBox14.Text)
                                gr1.gdireccion_partida = Trim(TextBox3.Text)
                                gr1.gNOTASALIDA = DateTime.Now.ToString("yyyy") & Label23.Text & "2" & Label3.Text
                                gr1.gchofer = Trim(TextBox16.Text)
                                gr1.gserie = Trim(TextBox7.Text)
                                gr1.gcorrelativo = Trim(TextBox17.Text)
                                gr1.galmacen = Trim(Label23.Text)
                                gr1.glicencia = Trim(TextBox15.Text)
                                gr1.gusuario = "almace"
                                gr1.gfase = ""
                                gr1.gdestino = ComboBox2.SelectedValue.ToString
                                gr1.gtrpo = Trim(TextBox4.Text)
                                gr1.gccia = Trim(Label30.Text)
                                gr1.gglosa = Trim(TextBox6.Text)
                                gr1.gordens_3 = Trim(TextBox20.Text)
                                gr1.gsubfase_3 = Trim(TextBox24.Text)
                                If RadioButton4.Checked = True Then
                                    gr1.gsituacion = "1"
                                Else
                                    If RadioButton3.Checked = True Then
                                        gr1.gsituacion = "2"
                                    End If

                                End If

                                gr1.gmotivo = ComboBox1.SelectedValue.ToString

                                If ns2.insertar_cabecera_nsa(ns1) = True And gr2.insertar_cabecera_guia_sistema(gr1) = True Then
                                    MessageBox.Show("SE REGISTRO LA GUIA REMISION CORRECTAMENTE ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)

                                    Dim i10, a As Integer
                                    i10 = DataGridView1.Rows.Count

                                    For a = 0 To i10 - 1

                                        Dim ns3 As New vnsadetalle
                                        Dim ns4 As New fnsa
                                        Dim nu2 As String

                                        nu2 = DataGridView1.Rows(a).Cells(0).Value
                                        Select Case nu2.Length
                                            Case "1" : nu2 = "00" & "" & nu2
                                            Case "2" : nu2 = "0" & "" & nu2
                                            Case "3" : nu2 = nu2
                                        End Select

                                        ns3.gncom = Label3.Text
                                        ns3.gitem = nu2
                                        ns3.glinea = DataGridView1.Rows(a).Cells(2).Value
                                        ns3.gcantidad = DataGridView1.Rows(a).Cells(7).Value
                                        ns3.gund = DataGridView1.Rows(a).Cells(8).Value
                                        If Trim(DataGridView1.Rows(a).Cells(1).Value) = "" Then
                                            ns3.gop = ""
                                        Else
                                            ns3.gop = DataGridView1.Rows(a).Cells(1).Value
                                        End If
                                        ns3.gunidad_medidad = DataGridView1.Rows(a).Cells(5).Value
                                        ns3.gcantenvio = 0
                                        If Convert.ToString(DataGridView1.Rows(a).Cells(10).Value) = "" Then
                                            ns3.gpartida = ""
                                        Else
                                            ns3.gpartida = DataGridView1.Rows(a).Cells(10).Value
                                        End If

                                        ns3.grollo1 = DataGridView1.Rows(a).Cells(4).Value
                                        ns3.galmacen = Label23.Text
                                        If Convert.ToString(DataGridView1.Rows(a).Cells(9).Value) = "" Then
                                            ns3.gordtejeduria = ""
                                        Else
                                            ns3.gordtejeduria = DataGridView1.Rows(a).Cells(9).Value
                                        End If
                                        ns3.gano = Label29.Text
                                        If Convert.ToString(DataGridView1.Rows(a).Cells(10).Value) = "" Then
                                            ns3.glote = ""
                                        Else
                                            ns3.glote = DataGridView1.Rows(a).Cells(10).Value
                                        End If

                                        Select Case Trim(DataGridView1.Rows(a).Cells(22).Value)
                                            Case "VERDE" : ns3.gclasif = "1"
                                            Case "AMARILLO" : ns3.gclasif = "2"
                                            Case "CELESTE" : ns3.gclasif = "3"
                                            Case "ROJO" : ns3.gclasif = "4"
                                            Case "" : ns3.gclasif = ""
                                        End Select
                                        ns3.gccia = Label30.Text
                                        ns3.gOcompra = ""
                                        ns4.insertar_detalle_nsa(ns3)

                                        '-------------------------------- etalle guia
                                        Dim gr3 As New vguiadetalle
                                        Dim gr4 As New fguiasistema
                                        Dim nu1 As String
                                        Dim jj As New fingresopac
                                        Dim aa As New vpackingtela
                                        gr3.gsfactu = TextBox1.Text
                                        gr3.gnfactu = TextBox2.Text
                                        nu1 = DataGridView1.Rows(a).Cells(0).Value
                                        Select Case nu1.Length
                                            Case "1" : nu1 = "00" & "" & nu1
                                            Case "2" : nu1 = "0" & "" & nu1
                                            Case "3" : nu1 = nu1
                                        End Select
                                        gr3.gitems = nu1
                                        gr3.glinea = DataGridView1.Rows(a).Cells(2).Value
                                        gr3.gcantidad = DataGridView1.Rows(a).Cells(7).Value
                                        gr3.gcanencio = 0
                                        If Trim(DataGridView1.Rows(a).Cells(1).Value).Length = 0 Then
                                            gr3.gordens = ""
                                        Else
                                            gr3.gordens = DataGridView1.Rows(a).Cells(1).Value
                                        End If

                                        gr3.grollo = DataGridView1.Rows(a).Cells(4).Value
                                        If Convert.ToString(DataGridView1.Rows(a).Cells(10).Value) = "" Then
                                            gr3.gparti = ""
                                        Else
                                            gr3.gparti = DataGridView1.Rows(a).Cells(10).Value
                                        End If

                                        gr3.gunidad_medida = DataGridView1.Rows(a).Cells(5).Value
                                        gr3.gunidad_medida2 = DataGridView1.Rows(a).Cells(8).Value
                                        'gr3.gordens = ""
                                        If Convert.ToString(DataGridView1.Rows(a).Cells(6).Value) = "" Then
                                            gr3.glote = ""
                                        Else
                                            gr3.glote = DataGridView1.Rows(a).Cells(6).Value
                                        End If
                                        If Convert.ToString(DataGridView1.Rows(a).Cells(9).Value) = "" Then
                                            gr3.gordtejeduria = ""
                                        Else
                                            gr3.gordtejeduria = DataGridView1.Rows(a).Cells(9).Value
                                        End If
                                        gr3.gccia = Label30.Text

                                        If Convert.ToString(DataGridView1.Rows(a).Cells(17).Value) = "" Then
                                            gr3.gobser1_3a = ""
                                        Else
                                            gr3.gobser1_3a = DataGridView1.Rows(a).Cells(17).Value
                                        End If
                                        If Convert.ToString(DataGridView1.Rows(a).Cells(18).Value) = "" Then
                                            gr3.gobser2_3a = ""
                                        Else
                                            gr3.gobser2_3a = DataGridView1.Rows(a).Cells(18).Value
                                        End If
                                        If Convert.ToString(DataGridView1.Rows(a).Cells(19).Value) = "" Then
                                            gr3.gobser3_3a = ""
                                        Else
                                            gr3.gobser3_3a = DataGridView1.Rows(a).Cells(19).Value
                                        End If
                                        If Convert.ToString(DataGridView1.Rows(a).Cells(20).Value) = "" Then
                                            gr3.gobser4_3a = ""
                                        Else
                                            gr3.gobser4_3a = DataGridView1.Rows(a).Cells(20).Value
                                        End If
                                        If Convert.ToString(DataGridView1.Rows(a).Cells(21).Value) = "" Then
                                            gr3.gobser5_3a = ""
                                        Else
                                            gr3.gobser5_3a = DataGridView1.Rows(a).Cells(21).Value
                                        End If
                                        Select Case Trim(DataGridView1.Rows(a).Cells(22).Value)
                                            Case "VERDE" : gr3.gclasif = "1"
                                            Case "AMARILLO" : gr3.gclasif = "2"
                                            Case "CELESTE" : gr3.gclasif = "3"
                                            Case "ROJO" : gr3.gclasif = "4"
                                            Case "" : gr3.gclasif = ""
                                        End Select

                                        gr4.insertar_detalle_guia_sistema(gr3)


                                        If Trim(DataGridView1.Rows(a).Cells(6).Value) = "" Then
                                            aa.gpartida = ""
                                        Else
                                            aa.gpartida = Trim(DataGridView1.Rows(a).Cells(6).Value)
                                        End If
                                        aa.galmacen = Label23.Text
                                        aa.gseleccionado = 2
                                        aa.gguia = Trim(TextBox1.Text) & Trim(TextBox2.Text)

                                        jj.actualizar_packing(aa)

                                        If Trim(DataGridView1.Rows(a).Cells(14).Value) <> "" Then
                                            abrir()
                                            Dim cmd157 As New SqlCommand("update requerimiento_avios_detalle set estado =2 where id_requeminieto_detalle=@id_requeminieto_detalle", conx)
                                            cmd157.Parameters.AddWithValue("@id_requeminieto_detalle", Trim(DataGridView1.Rows(a).Cells(14).Value))
                                            cmd157.ExecuteNonQuery()
                                            Dim a11, b As Integer

                                            Dim sql102213 As String = "select COUNT(id_requeminieto_detalle) from requerimiento_avios_detalle where id_cabecera ='" + Trim(DataGridView1.Rows(a).Cells(15).Value) + "' and estado ='2'"
                                            Dim cmd102213 As New SqlCommand(sql102213, conx)
                                            Rsr3 = cmd102213.ExecuteReader()
                                            If Rsr3.Read() = True Then
                                                a11 = Rsr3(0)
                                            End If
                                            Rsr3.Close()
                                            Dim sql1022134 As String = "select COUNT(id_requeminieto_detalle) from requerimiento_avios_detalle where id_cabecera ='" + Trim(DataGridView1.Rows(a).Cells(15).Value) + "'"
                                            Dim cmd1022134 As New SqlCommand(sql1022134, conx)
                                            Rsr21 = cmd1022134.ExecuteReader()
                                            If Rsr21.Read() = True Then
                                                b = Rsr21(0)
                                            End If
                                            Rsr21.Close()

                                            If a11 = b Then
                                                Dim cmd1577 As New SqlCommand("update requerimiento_avios_cabecera set estado = 4 where id_requeminieto='" + Trim(DataGridView1.Rows(a).Cells(15).Value) + "'", conx)
                                                cmd1577.Parameters.AddWithValue("@id_requeminieto_detalle", Trim(DataGridView1.Rows(a).Cells(14).Value))
                                                cmd1577.ExecuteNonQuery()
                                            End If
                                        End If
                                        If Label35.Text = "1" Then
                                            Dim sum, suma As Double
                                            Dim cmd20 As New SqlCommand("update DetalleTelaManufactura  set EstDtm = '2'  where OpDtm =@op and CopDtm =@linea and CodDtm =@codigo", conx)
                                            cmd20.Parameters.AddWithValue("@op", (DataGridView1.Rows(a).Cells(1).Value.ToString.Trim))
                                            cmd20.Parameters.AddWithValue("@linea", (DataGridView1.Rows(a).Cells(2).Value.ToString.Trim))
                                            cmd20.Parameters.AddWithValue("@codigo", (DataGridView1.Rows(a).Cells(23).Value.ToString.Trim))
                                            cmd20.ExecuteNonQuery()

                                            Dim sql10 As String = "select ROUND(CASE WHEN cg.familia_17 = 'TPL' THEN cp.cxplineal * cantp_3 ELSE cp.kilos * q.cantp_3 END, 2) 
                                                                        from custom_vianny.dbo.Consumo_Op_DET cp
                                                                        LEFT JOIN custom_vianny.dbo.cag1700 cg ON cp.ccia = cg.ccia AND cp.tela = cg.linea_17 
                                                                        LEFT JOIN custom_vianny.dbo.qag0300 q ON cp.ccia = q.ccia AND cp.op = q.ncom_3 
                                                                        where op ='" + Trim(DataGridView1.Rows(a).Cells(1).Value.ToString.Trim) + "' and cp.ccia ='" + Label30.Text + "'"
                                            Dim cmd10 As New SqlCommand(sql10, conx)
                                            Rsr215 = cmd10.ExecuteReader()
                                            If Rsr215.Read() = True Then
                                                sum = Rsr215(0)
                                            End If
                                            Rsr215.Close()

                                            Dim sql101 As String = "select sum(CanDtm) from DetalleTelaManufactura WHERE  OpDtm ='" + DataGridView1.Rows(a).Cells(1).Value.ToString.Trim + "'"
                                            Dim cmd101 As New SqlCommand(sql101, conx)
                                            Rsr216 = cmd101.ExecuteReader()
                                            If Rsr216.Read() = True Then
                                                suma = Rsr216(0)
                                            End If
                                            Rsr216.Close()

                                            'If sum = suma Then

                                            '    Dim cmd21 As New SqlCommand("update RequerimientoTelaProd set EstRtp ='2' where IdRtp in (@IdRtp)", conx)
                                            '    cmd21.Parameters.AddWithValue("@IdRtp", Trim(Label36.Text))
                                            '    cmd21.ExecuteNonQuery()
                                            'Else
                                            '    Dim cmd211 As New SqlCommand("update RequerimientoTelaProd set EstRtp ='1' where IdRtp in (@IdRtp)", conx)
                                            '    cmd211.Parameters.AddWithValue("@IdRtp", Trim(Label36.Text))
                                            '    cmd211.ExecuteNonQuery()
                                            'End If
                                            ReqTelaProduccion.Button1.PerformClick()
                                        End If

                                    Next
                                    'End If
                                    Dim agregar As String = "delete from Almacen_Ficticio where  AlmAfi='" + Label23.Text + "' and EmpAfi='" + Label30.Text + "'  and GuiAfi ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' "
                                    ELIMINAR(agregar)
                                    If Label23.Text = "37" Or Label23.Text = "38" Then
                                        Dim jj As Integer
                                        Dim jlk As New fpedidopacking
                                        Dim kll As New vpackign
                                        Dim kll1 As New vpackign
                                        jj = DataGridView5.Rows.Count
                                        For i = 0 To jj - 1
                                            kll.gpartida = DataGridView5.Rows(i).Cells(1).Value
                                            kll.gID = DataGridView5.Rows(i).Cells(0).Value
                                            kll.gselec = 1
                                            kll.galmacen = DataGridView5.Rows(i).Cells(3).Value
                                            kll.gestado = 0
                                            jlk.actualizar_rollos(kll)
                                            kll1.galmacen = DataGridView5.Rows(i).Cells(3).Value
                                            kll1.gguia = TextBox1.Text & "-" & TextBox2.Text
                                            kll1.gID = DataGridView5.Rows(i).Cells(0).Value
                                            kll1.gpartida = DataGridView5.Rows(i).Cells(1).Value
                                            jlk.actualizar_guia(kll1)
                                        Next

                                    End If

                                    If Trim(ComboBox1.Text) = "VENTA" Then
                                        'ENVIAR_CORREO()
                                    End If
                                    Dim hj2 As New flog
                                    Dim hj1 As New vlog
                                    hj1.gmodulo = "NSA-ALMACEN"
                                    hj1.gcuser = MDIParent1.Label3.Text
                                    If Label32.Text = "0" Then
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
                                    hj12.gmodulo = "GUIA-ALMACEN"
                                    hj12.gcuser = MDIParent1.Label3.Text

                                    If Label32.Text = "0" Then
                                        hj12.gaccion = 1
                                    Else
                                        hj12.gaccion = 2
                                    End If
                                    hj12.gpc = My.Computer.Name
                                    hj12.gfecha = DateTimePicker3.Value
                                    hj12.gdetalle = Trim(Label23.Text) & Trim(TextBox1.Text) & Trim(TextBox2.Text)
                                    hj12.gccia = Label30.Text
                                    hj22.insertar_log(hj12)

                                    Dim respuesta As DialogResult

                                    respuesta = MessageBox.Show("DESEA IMPRIMIR EL PACKING?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                    If (respuesta = Windows.Forms.DialogResult.Yes) Then

                                        'Reporte_Guia.TextBox1.Text = TextBox1.Text
                                        'Reporte_Guia.TextBox2.Text = TextBox2.Text
                                        'Reporte_Guia.TextBox3.Text = Label30.Text
                                        'Reporte_Guia.Show()

                                        'Packing_nacional.TextBox1.Text = DataGridView1.Rows(num1).Cells(1).Value
                                        '    Packing_nacional.TextBox2.Text = DataGridView1.Rows(num1).Cells(3).Value
                                        '    Packing_nacional.TextBox3.Text = "68"
                                        '    Packing_nacional.Show()

                                    End If
                                    'respuesta1 = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                    'If (respuesta1 = Windows.Forms.DialogResult.Yes) Then

                                    '    'Rpt_Packing.TextBox1.Text = Trim(TextBox1.Text) & "-" & Trim(TextBox2.Text)
                                    '    'Rpt_Packing.Show()



                                    Dim func As New fguiasistema
                                    Dim func1 As New vguiacabecera

                                    func1.gserie = TextBox1.Text
                                    func1.galmacen = Label23.Text
                                    func1.gccia = Label30.Text
                                    TextBox2.Text = func.mostrar_guia_correlativo(func1) + 1

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
                                    func1.gmes = DateTime.Now.ToString("MM")
                                    func1.galmacen = Label23.Text
                                    func1.gano = Label29.Text
                                    func1.gccia = Label30.Text
                                    Label3.Text = func.mostrar_correlativo_nsa(func1) + 1
                                    Select Case Label3.Text.Length

                                        Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
                                        Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
                                        Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
                                        Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
                                    End Select
                                    RadioButton1.Checked = True
                                    Button1.Enabled = False
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
                                    TextBox18.Text = ""
                                    TextBox7.Text = ""
                                    TextBox17.Text = ""
                                    Label35.Text = "0"
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
                                    ComboBox3.Enabled = False
                                    PictureBox1.Enabled = False
                                    PictureBox2.Enabled = False
                                    Button5.Enabled = False
                                    PictureBox4.Enabled = False
                                    PictureBox5.Enabled = False
                                    DataGridView5.Rows.Clear()
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Dim Rsr3, Rsr21 As SqlDataReader
    Private Sub Guia_remision_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Dim j As Integer
        Dim jlk As New fpedidopacking
        Dim kll As New vpackign
        j = DataGridView5.Rows.Count
        For i = 0 To j - 1
            kll.gpartida = DataGridView5.Rows(i).Cells(1).Value
            kll.gID = DataGridView5.Rows(i).Cells(0).Value
            kll.gselec = 0
            kll.galmacen = DataGridView5.Rows(i).Cells(3).Value
            kll.gestado = 1
            jlk.actualizar_rollos(kll)
        Next
        Dim agregar As String = "delete from Almacen_Ficticio where  AlmAfi='" + Label23.Text + "' and EmpAfi='" + Label30.Text + "'  and GuiAfi ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' "
        ELIMINAR(agregar)

    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim fu As New vguiacabecera
            Dim func As New fguiasistema
            Dim fu1 As New vguiacabecera

            Dim func1 As New vguiacabecera
            Select Case TextBox1.Text.Length

                Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = TextBox1.Text

            End Select
            TextBox1.Text = TextBox1.Text
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
            func1.gmes = DateTime.Now.ToString("MM")
            func1.galmacen = Label23.Text
            func1.gano = Label29.Text
            func1.gccia = Label30.Text
            Label3.Text = func.mostrar_correlativo_nsa(func1) + 1
            Select Case Label3.Text.Length

                Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
                Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
                Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
                Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
            End Select
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
            TextBox4.Text = ""
            TextBox7.Text = ""
            TextBox17.Text = ""
            TextBox5.Text = ""
            TextBox14.Text = ""
            TextBox15.Text = ""
            TextBox16.Text = ""
            TextBox6.Text = ""
            TextBox18.Text = ""
            TextBox7.Text = ""
            TextBox17.Text = ""
            TextBox11.Text = "20508740361"
            TextBox12.Text = "CONSORCIO TEXTIL VIANNY SAC"
            TextBox13.Text = "JR. LOS OLMOS NRO. 358 URB. CANTO BELLO (ALT. PARDERO 1 AV."
            DataGridView1.Rows.Clear()
            TextBox2.Select()
        End If
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged

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

        Dim func1 As New vguiacabecera

        func1.gserie = TextBox1.Text
        func1.galmacen = Label23.Text
        func1.gccia = Label30.Text
        TextBox2.Text = func.mostrar_guia_correlativo(func1) + 1

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
        func1.gmes = DateTime.Now.ToString("MM")
        func1.galmacen = Label23.Text
        func1.gano = Label29.Text
        func1.gccia = Label30.Text
        Label3.Text = func.mostrar_correlativo_nsa(func1) + 1
        Select Case Label3.Text.Length

            Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
            Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
            Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
            Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
        End Select
        RadioButton1.Checked = True
        Button1.Enabled = False
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
        ComboBox3.Enabled = False
        PictureBox1.Enabled = False
        PictureBox2.Enabled = False
        Button5.Enabled = False
        PictureBox4.Enabled = False
        PictureBox5.Enabled = False
        DataGridView5.Rows.Clear()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox3.Text = "AV. CANTO BELLO 152 - URB. CANTO GRANDE SJL"
        Else
            TextBox3.Text = "MZ.E Lote, PARCELACION RUSTICA HUERTOS DE ORO DE SAN HILARION Y PAMPA Y HOYADAS DE CALANGUILLO - CHILCA - CAÑETE -LIMA"
        End If
    End Sub


    Private Function IsValidNumber(value As String) As Boolean
        Dim pattern As String = "^\d+(\.\d+)?$"
        Return System.Text.RegularExpressions.Regex.IsMatch(value, pattern)
    End Function

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        If e.ColumnIndex = 4 Or e.ColumnIndex = 7 Then
            Dim cellValue As String = e.FormattedValue.ToString()
            If Not IsValidNumber(cellValue) Then
                MessageBox.Show("El valor ingresado no es válido. Debe ser un número entero o decimal con al menos un número antes del punto decimal.", "Error de Validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
            End If
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Trim(TextBox8.Text) = "" Then
            MsgBox("NO EXISTE NINGUN RUC, INGRESELO PARA UTILIZAR ESTA OPCION")
        Else
            Agregar_requerimiento.TextBox1.Text = Trim(TextBox10.Text)
            Agregar_requerimiento.Label2.Text = TextBox8.Text
            Agregar_requerimiento.Label5.Text = ComboBox3.Text
            Agregar_requerimiento.Label3.Text = Label13.Text
            Agregar_requerimiento.Label4.Text = Label10.Text
            Agregar_requerimiento.Label6.Text = 2
            Agregar_requerimiento.Show()
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        log.Label1.Text = Trim(Label23.Text) & Trim(TextBox1.Text) & Trim(TextBox2.Text)
        log.Label2.Text = "GUIA-ALMACEN"
        log.Label3.Text = Label30.Text
        log.Show()
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click

    End Sub

    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs)

    End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.F1 Then
            proceso.Label2.Text = 2
            proceso.ShowDialog()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OBS" Then
            'OBSERVACION.Label2.Text = 5
            'OBSERVACION.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(17).Value)
            'OBSERVACION.Label3.Text = e.RowIndex
            'OBSERVACION.ShowDialog()

            OB_SC.Label1.Text = e.RowIndex
            OB_SC.Label2.Text = 2
            OB_SC.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(17).Value)
            OB_SC.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(18).Value)
            OB_SC.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(19).Value)
            OB_SC.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(20).Value)
            OB_SC.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(21).Value)
            OB_SC.ShowDialog()
        End If
    End Sub

    Private Sub TextBox2_QueryAccessibilityHelp(sender As Object, e As QueryAccessibilityHelpEventArgs) Handles TextBox2.QueryAccessibilityHelp

    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim columnIndex2 As Integer = 1
        Dim columnIndex As Integer = 4
        Dim columnIndex1 As Integer = 4
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex AndAlso TypeOf e.Control Is TextBox Then
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex1 AndAlso TypeOf e.Control Is TextBox Then
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex2 AndAlso TypeOf e.Control Is TextBox Then
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If

        Dim txtBox As TextBox = TryCast(e.Control, TextBox)

        If txtBox IsNot Nothing Then
            ' Establece el CharacterCasing a Upper para convertir a mayúsculas automáticamente
            txtBox.CharacterCasing = CharacterCasing.Upper
        End If
    End Sub

    Private Sub DataGridView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DataGridView1.KeyPress
        Dim currentCell As DataGridViewCell = DataGridView1.CurrentCell
        If currentCell.ColumnIndex = 7 Then
            If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back AndAlso e.KeyChar <> "." Then
                ' Si el carácter no es numérico, de retroceso ni un punto, cancela el evento KeyPress
                e.Handled = True
            End If
            If e.KeyChar = "." AndAlso DirectCast(sender, TextBox).Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        End If
        If currentCell.ColumnIndex = 4 Then
            If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
                ' Si el carácter no es numérico ni de retroceso, cancela el evento KeyPress
                e.Handled = True
            End If
        End If
        If currentCell.ColumnIndex = 1 Then
            If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
                ' Si el carácter no es numérico ni de retroceso, cancela el evento KeyPress
                e.Handled = True
            End If
        End If
    End Sub
End Class