Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class Urdido
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2 As New DataTable
    Public enunciado4 As SqlCommand
    Public respuesta4 As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public enunciado5 As SqlCommand
    Public respuesta5 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Urdido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim func As New furdido
        Me.TextBox1.Text = func.buscar_correla_urdido + 1
        Select Case TextBox1.Text.Length
            Case "1" : TextBox1.Text = "000000000" & "" & TextBox1.Text
            Case "2" : TextBox1.Text = "00000000" & "" & TextBox1.Text
            Case "3" : TextBox1.Text = "0000000" & "" & TextBox1.Text
            Case "4" : TextBox1.Text = "000000" & "" & TextBox1.Text
            Case "5" : TextBox1.Text = "00000" & "" & TextBox1.Text
            Case "6" : TextBox1.Text = "0000" & "" & TextBox1.Text
            Case "7" : TextBox1.Text = "000" & "" & TextBox1.Text
            Case "8" : TextBox1.Text = "00" & "" & TextBox1.Text
            Case "9" : TextBox1.Text = "0" & "" & TextBox1.Text
            Case "10" : TextBox1.Text = TextBox1.Text
        End Select
        ComboBox2.SelectedIndex = 0
        PictureBox2.Enabled = False
        PictureBox3.Enabled = False
        PictureBox4.Enabled = False
        Button2.Enabled = False
        TextBox10.Select()
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs)

    End Sub

    'Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        Dim fg As New vpartida
    '        Dim fg1 As New fpartida
    '        Select Case TextBox11.Text.Length
    '            Case "1" : TextBox11.Text = "0000" & "" & TextBox11.Text
    '            Case "2" : TextBox11.Text = "000" & "" & TextBox11.Text
    '            Case "3" : TextBox11.Text = "00" & "" & TextBox11.Text
    '            Case "4" : TextBox11.Text = "0" & "" & TextBox11.Text
    '            Case "5" : TextBox11.Text = TextBox11.Text
    '        End Select
    '        Select Case TextBox11.Text.Length
    '            Case "1" : fg.gcodigo = "0000" & "" & TextBox11.Text
    '            Case "2" : fg.gcodigo = "000" & "" & TextBox11.Text
    '            Case "3" : fg.gcodigo = "00" & "" & TextBox11.Text
    '            Case "4" : fg.gcodigo = "0" & "" & TextBox11.Text
    '            Case "5" : fg.gcodigo = TextBox11.Text
    '        End Select

    '        TextBox12.Text = fg1.buscar_persona(fg)
    '    End If

    'End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add(1)
        Dim ui As Integer
        ui = DataGridView1.Rows.Count

        If ui = 1 Then
            DataGridView1.Rows(0).Cells(0).Value = 1
            DataGridView1.Rows(0).Cells(2).Value = DateTime.Now.ToString("yyyy/MM/dd") & " " & DateTime.Now.ToString("hh:mm:ss")
        Else
            DataGridView1.Rows(ui - 1).Cells(0).Value = DataGridView1.Rows(ui - 2).Cells(0).Value + 1
            DataGridView1.Rows(ui - 1).Cells(2).Value = DateTime.Now.ToString("yyyy/MM/dd") & " " & DateTime.Now.ToString("hh:mm:ss")

        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        DataGridView1.DataSource = ""
        ComboBox1.Items.Clear()
        TextBox2.Text = ""
        TextBox3.Text = ""

        TextBox5.Text = ""

        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox4.Text = ""

    End Sub
    Sub llenar_combo_box2(ByVal cb As ComboBox, ByVal alm As String)
        Try

            conn = New SqlDataAdapter("select urdido from programa_tenido_detalle where  programa =" + alm, conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "urdido"

            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            ty2.Clear()

            Select Case TextBox10.Text.Length
                Case "1" : TextBox10.Text = "000000000" & "" & TextBox10.Text
                Case "2" : TextBox10.Text = "00000000" & "" & TextBox10.Text
                Case "3" : TextBox10.Text = "0000000" & "" & TextBox10.Text
                Case "4" : TextBox10.Text = "000000" & "" & TextBox10.Text
                Case "5" : TextBox10.Text = "00000" & "" & TextBox10.Text
                Case "6" : TextBox10.Text = "0000" & "" & TextBox10.Text
                Case "7" : TextBox10.Text = "000" & "" & TextBox10.Text
                Case "8" : TextBox10.Text = "00" & "" & TextBox10.Text
                Case "9" : TextBox10.Text = "0" & "" & TextBox10.Text
                Case "10" : TextBox10.Text = TextBox10.Text
            End Select

            abrir()
            llenar_combo_box2(ComboBox1, TextBox10.Text)
            Button2.Enabled = True
            PictureBox4.Enabled = True
            ComboBox1.SelectedIndex = -1
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim hj As New vurdido
        Dim hj1 As New vurdido
        Dim hj2 As New vurdido
        Dim gg As New furdido
        Dim ij As Integer

        'eliminar urdido
        hj2.gid = TextBox1.Text

        gg.eliminar_urdido(hj2)
        'insertar cabecera
        hj.gid = TextBox1.Text
        hj.gjuego_urdido = ComboBox1.Text
        hj.garticulo = TextBox2.Text
        hj.gtitulo = TextBox3.Text
        hj.ghilo = TextBox5.Text
        hj.glongitud = TextBox7.Text
        hj.glote = TextBox8.Text
        hj.gball = TextBox9.Text

        hj.gobservacion = TextBox4.Text
        hj.gflag = 1
        hj.gprograma = TextBox10.Text
        hj.gmaquina = ComboBox2.Text
        hj.gfecha = DateTimePicker1.Value
        ' insertar detalle
        ij = DataGridView1.Rows.Count
        If gg.insertar_urdido_cabecera(hj) = True Then
            For i = 0 To ij - 1
                hj1.gid = TextBox1.Text
                hj1.gjuego_urdidod = ComboBox1.Text
                If Convert.ToString(DataGridView1.Rows(i).Cells(0).Value) = "" Then
                    hj1.gitemsd = ""
                Else
                    hj1.gitemsd = DataGridView1.Rows(i).Cells(0).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(1).Value) = "" Then
                    hj1.gballd = ""
                Else
                    hj1.gballd = DataGridView1.Rows(i).Cells(1).Value
                End If

                hj1.gfechad = DataGridView1.Rows(i).Cells(2).Value
                If Convert.ToString(DataGridView1.Rows(i).Cells(3).Value) = "" Then
                    hj1.gcodigod = ""
                Else
                    hj1.gcodigod = DataGridView1.Rows(i).Cells(3).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                    hj1.goperariod = ""
                Else
                    hj1.goperariod = DataGridView1.Rows(i).Cells(4).Value
                End If


                hj1.gturno1d = DataGridView1.Rows(i).Cells(5).Value
                hj1.gturno2d = DataGridView1.Rows(i).Cells(6).Value
                hj1.gturno3d = DataGridView1.Rows(i).Cells(7).Value  ' MAQUINA
                If Convert.ToString(DataGridView1.Rows(i).Cells(8).Value) = "" Then
                    hj1.ghrotosd = ""
                Else
                    hj1.ghrotosd = DataGridView1.Rows(i).Cells(8).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(9).Value) = "" Then
                    hj1.ginciod = ""
                Else
                    hj1.ginciod = DataGridView1.Rows(i).Cells(9).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(10).Value) = "" Then
                    hj1.gfind = ""
                Else
                    hj1.gfind = DataGridView1.Rows(i).Cells(10).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(12).Value) = "" Then ' PESO BALL
                    hj1.gbolsasd = ""
                Else
                    hj1.gbolsasd = DataGridView1.Rows(i).Cells(12).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(11).Value) = "" Then
                    hj1.gconosd = ""
                Else
                    hj1.gconosd = DataGridView1.Rows(i).Cells(11).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(13).Value) = "" Then ' SALDO PESO
                    hj1.gayudanted = ""
                Else
                    hj1.gayudanted = DataGridView1.Rows(i).Cells(13).Value
                End If

                hj1.gflagd = 1
                hj1.gprogramad = TextBox10.Text
                gg.insertar_urdido_detalle(hj1)
            Next
            MsgBox("LA INFORMACION SE REGISTRO CON EXITO")
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""

            TextBox4.Text = ""
            TextBox10.Text = ""
            DataGridView1.Rows.Clear()
            Dim func As New furdido
            Me.TextBox1.Text = func.buscar_correla_urdido + 1
            Select Case TextBox1.Text.Length
                Case "1" : TextBox1.Text = "000000000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "00000000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0000000" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "000000" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = "00000" & "" & TextBox1.Text
                Case "6" : TextBox1.Text = "0000" & "" & TextBox1.Text
                Case "7" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "8" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "9" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "10" : TextBox1.Text = TextBox1.Text
            End Select
            TextBox10.Select()

        End If

    End Sub



    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        abrir()
        enunciado2 = New SqlCommand("select articulo from programa_tenido_detalle where urdido ='" + ComboBox1.Text + "'", conx)
        respuesta2 = enunciado2.ExecuteReader
        While respuesta2.Read
            TextBox2.Text = respuesta2.Item("articulo")
        End While
        respuesta2.Close()

        enunciado4 = New SqlCommand("select titulo from programa_tenido_detalle where urdido ='" + Trim(ComboBox1.Text) + "'", conx)
        respuesta4 = enunciado4.ExecuteReader
        While respuesta4.Read
            TextBox3.Text = respuesta4.Item("titulo")
        End While
        respuesta4.Close()

        enunciado5 = New SqlCommand("select hilo from programa_tenido_detalle where urdido ='" + Trim(ComboBox1.Text) + "'", conx)
        respuesta5 = enunciado5.ExecuteReader
        While respuesta5.Read
            TextBox5.Text = respuesta5.Item("hilo")
        End While
        respuesta5.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        DataGridView1.BeginEdit(True)
        Dim J As Integer

        J = DataGridView1.Rows.Count
        For I = 0 To J - 1
            If Me.DataGridView1.CurrentRow.Index.ToString() = I Then
                DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Yellow
            Else
                DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.White
            End If
        Next
    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        If TypeOf e.Control Is TextBox Then

            DirectCast(e.Control, TextBox).CharacterCasing = CharacterCasing.Upper

        End If
    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs)
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Dim kj As New DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim hj As New furdido
            Dim jka As New vurdido
            Dim ll As Integer


            Select Case TextBox1.Text.Length
                Case "1" : TextBox1.Text = "000000000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "00000000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0000000" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "000000" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = "00000" & "" & TextBox1.Text
                Case "6" : TextBox1.Text = "0000" & "" & TextBox1.Text
                Case "7" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "8" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "9" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "10" : TextBox1.Text = TextBox1.Text
            End Select
            jka.gid = TextBox1.Text
            kj = hj.buscar_urdido(jka)
            DataGridView1.Rows.Clear()
            If kj.Rows.Count <> 0 Then

                DataGridView2.DataSource = kj

                ll = DataGridView2.Rows.Count
                DataGridView1.Rows.Add(ll)

                For i = 0 To ll - 1

                    TextBox2.Text = DataGridView2.Rows(0).Cells(2).Value
                    TextBox3.Text = DataGridView2.Rows(0).Cells(3).Value
                    TextBox5.Text = DataGridView2.Rows(0).Cells(4).Value
                    TextBox7.Text = DataGridView2.Rows(0).Cells(5).Value
                    TextBox8.Text = DataGridView2.Rows(0).Cells(6).Value
                    TextBox9.Text = DataGridView2.Rows(0).Cells(7).Value

                    TextBox4.Text = DataGridView2.Rows(0).Cells(8).Value
                    TextBox10.Text = DataGridView2.Rows(0).Cells(10).Value

                    DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(16).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(i).Cells(17).Value
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView2.Rows(i).Cells(18).Value
                    DataGridView1.Rows(i).Cells(3).Value = DataGridView2.Rows(i).Cells(19).Value
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(20).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(21).Value
                    DataGridView1.Rows(i).Cells(6).Value = DataGridView2.Rows(i).Cells(22).Value
                    DataGridView1.Rows(i).Cells(7).Value = DataGridView2.Rows(i).Cells(23).Value
                    DataGridView1.Rows(i).Cells(8).Value = DataGridView2.Rows(i).Cells(24).Value
                    DataGridView1.Rows(i).Cells(9).Value = DataGridView2.Rows(i).Cells(25).Value
                    DataGridView1.Rows(i).Cells(10).Value = DataGridView2.Rows(i).Cells(26).Value
                    DataGridView1.Rows(i).Cells(11).Value = DataGridView2.Rows(i).Cells(28).Value
                    DataGridView1.Rows(i).Cells(12).Value = DataGridView2.Rows(i).Cells(27).Value
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView2.Rows(i).Cells(29).Value
                Next
                abrir()
                llenar_combo_box2(ComboBox1, TextBox10.Text)
                ComboBox1.Text = Trim(DataGridView2.Rows(0).Cells(1).Value).ToString
                Select Case ComboBox2.Text
                    Case "MAQ 1" : ComboBox2.SelectedIndex = 0
                    Case "MAQ 2" : ComboBox2.SelectedIndex = 1
                    Case "MAQ 3" : ComboBox2.SelectedIndex = 2

                End Select
                TextBox2.Enabled = False
                TextBox3.Enabled = False
                TextBox5.Enabled = False
                TextBox7.Enabled = False
                TextBox8.Enabled = False
                TextBox9.Enabled = False
                TextBox4.Enabled = False

                TextBox10.Enabled = False
                DataGridView1.Enabled = False
                PictureBox2.Enabled = True
                PictureBox3.Enabled = True
            Else
                DataGridView1.Rows.Clear()
                DataGridView2.DataSource = ""

                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox5.Text = ""
                TextBox7.Text = ""
                TextBox8.Text = ""
                TextBox9.Text = ""

                TextBox4.Text = ""
                TextBox10.Text = ""

            End If
            TextBox1.ReadOnly = True
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("QUIERES EDITAR ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then

            TextBox3.Enabled = True
            TextBox5.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            TextBox9.Enabled = True

            TextBox4.Enabled = True
            TextBox10.Enabled = True
            DataGridView1.Enabled = True
            Button2.Enabled = True

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
            I18 = DataGridView1.Rows.Count
            For i1 = 0 To I18 - 1
                VAL = i1 + 1

                DataGridView1.Rows(i1).Cells(0).Value = VAL

            Next
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, fila As Integer

        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CODIGO" Then
            Dim fg As New vpartida
            Dim fg1 As New fpartida
            If DataGridView1.Rows(fila).Cells(3).Value = "" Then
            Else
                Select Case DataGridView1.Rows(fila).Cells(3).Value.Length
                    Case "1" : DataGridView1.Rows(fila).Cells(3).Value = "0000" & "" & DataGridView1.Rows(fila).Cells(3).Value
                    Case "2" : DataGridView1.Rows(fila).Cells(3).Value = "000" & "" & DataGridView1.Rows(fila).Cells(3).Value
                    Case "3" : DataGridView1.Rows(fila).Cells(3).Value = "00" & "" & DataGridView1.Rows(fila).Cells(3).Value
                    Case "4" : DataGridView1.Rows(fila).Cells(3).Value = "0" & "" & DataGridView1.Rows(fila).Cells(3).Value
                    Case "5" : DataGridView1.Rows(fila).Cells(3).Value = DataGridView1.Rows(fila).Cells(3).Value
                End Select
                Select Case DataGridView1.Rows(fila).Cells(3).Value.Length
                    Case "1" : fg.gcodigo = "0000" & "" & DataGridView1.Rows(fila).Cells(3).Value
                    Case "2" : fg.gcodigo = "000" & "" & DataGridView1.Rows(fila).Cells(3).Value
                    Case "3" : fg.gcodigo = "00" & "" & DataGridView1.Rows(fila).Cells(3).Value
                    Case "4" : fg.gcodigo = "0" & "" & DataGridView1.Rows(fila).Cells(3).Value
                    Case "5" : fg.gcodigo = DataGridView1.Rows(fila).Cells(3).Value
                End Select


                DataGridView1.Rows(fila).Cells(4).Value = fg1.buscar_persona(fg)
            End If


        End If
    End Sub
End Class