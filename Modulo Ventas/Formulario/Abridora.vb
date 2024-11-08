Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class Abridora
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
    Public enunciado6 As SqlCommand
    Public respuesta6 As SqlDataReader
    Public enunciado7 As SqlCommand
    Public respuesta7 As SqlDataReader
    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            Dim fg As New vpartida
            Dim fg1 As New fpartida
            'Select Case TextBox14.Text.Length
            '    Case "1" : TextBox14.Text = "0000" & "" & TextBox14.Text
            '    Case "2" : TextBox14.Text = "000" & "" & TextBox14.Text
            '    Case "3" : TextBox14.Text = "00" & "" & TextBox14.Text
            '    Case "4" : TextBox14.Text = "0" & "" & TextBox14.Text
            '    Case "5" : TextBox14.Text = TextBox14.Text
            'End Select
            'Select Case TextBox14.Text.Length
            '    Case "1" : fg.gcodigo = "0000" & "" & TextBox14.Text
            '    Case "2" : fg.gcodigo = "000" & "" & TextBox14.Text
            '    Case "3" : fg.gcodigo = "00" & "" & TextBox14.Text
            '    Case "4" : fg.gcodigo = "0" & "" & TextBox14.Text
            '    Case "5" : fg.gcodigo = TextBox14.Text
            'End Select

            'TextBox13.Text = fg1.buscar_persona(fg)
        End If

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim jj As New vabridora
        Dim jj1 As New vabridora
        Dim jj2 As New vabridora
        Dim ll As New fabridora
        Dim j As Integer
        'eliminar abridora
        jj.gid = TextBox12.Text

        ll.eliminar_abridora(jj)

        'insertar cabecera
        jj1.gid = TextBox12.Text
        jj1.gjuego_abridora = TextBox1.Text
        'jj1.gcodigo = TextBox14.Text
        'jj1.gtrabajador = TextBox13.Text
        jj1.gjuego_urdido = ComboBox1.Text
        jj1.glongitud = TextBox7.Text
        jj1.gtitulo = TextBox3.Text
        jj1.ghilo = TextBox5.Text
        jj1.gjuego_tenido = TextBox11.Text
        jj1.glote = TextBox8.Text
        jj1.gflag = 1
        jj1.gball = TextBox9.Text
        jj1.gprograma = TextBox10.Text
        jj1.garticulo = TextBox2.Text
        jj1.gfecha = DateTimePicker1.Value
        If Trim(TextBox9.Text) = "" Then
            jj1.gjuego_conera = ""
        Else
            jj1.gjuego_conera = Trim(TextBox9.Text)
        End If

        If Trim(TextBox4.Text) = "" Then
            jj1.gcuerda = ""
        Else
            jj1.gcuerda = Trim(TextBox4.Text)
        End If

        'jj1.gmaquina = ComboBox2.Text

        j = DataGridView1.Rows.Count
        If ll.insertar_abridora_cabecera(jj1) = True Then
            For i = 0 To j - 1
                jj2.gid = TextBox12.Text
                jj2.gjuego_abridorad = TextBox1.Text
                jj2.gcuerdad = DataGridView1.Rows(i).Cells(0).Value
                jj2.gtachod = DataGridView1.Rows(i).Cells(1).Value
                jj2.gmaquinad = DataGridView1.Rows(i).Cells(2).Value
                jj2.gplegadord = DataGridView1.Rows(i).Cells(3).Value
                If DataGridView1.Rows(i).Cells(4).Value = "" Then
                    jj2.gfechad = DateTime.Now.ToString("yyyy/MM/dd")
                Else
                    jj2.gfechad = DataGridView1.Rows(i).Cells(4).Value
                End If

                jj2.gcodigo = DataGridView1.Rows(i).Cells(5).Value
                jj2.gtrabajador = DataGridView1.Rows(i).Cells(6).Value
                If DataGridView1.Rows(i).Cells(7).Value = "" Then
                    jj2.gtitulod = ""

                Else
                    jj2.gtitulod = DataGridView1.Rows(i).Cells(7).Value
                End If

                jj2.ghilod = DataGridView1.Rows(i).Cells(8).Value
                'jj2.gpesod = DataGridView1.Rows(i).Cells(6).Value
                If DataGridView1.Rows(i).Cells(9).Value = "" Then
                    jj2.gloted = ""
                Else
                    jj2.gloted = DataGridView1.Rows(i).Cells(9).Value
                End If
                If DataGridView1.Rows(i).Cells(10).Value = "" Then
                    jj2.gdespachod = ""
                Else
                    jj2.gdespachod = DataGridView1.Rows(i).Cells(10).Value
                End If

                jj2.gflagd = 1
                jj2.gprogramad = TextBox10.Text
                ll.insertar_abridora_detalle(jj2)
            Next
            MsgBox("LA INFORMACION SE REGISTRO CON EXITO")
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            TextBox7.Text = ""
            TextBox1.Text = ""
            TextBox11.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            'TextBox14.Text = ""
            'TextBox13.Text = ""
            TextBox10.Text = ""
            DataGridView1.Rows.Clear()
            Dim func As New fabridora
            Me.TextBox12.Text = func.correlativo_abridora + 1
            Select Case TextBox12.Text.Length
                Case "1" : TextBox12.Text = "000000000" & "" & TextBox12.Text
                Case "2" : TextBox12.Text = "00000000" & "" & TextBox12.Text
                Case "3" : TextBox12.Text = "0000000" & "" & TextBox12.Text
                Case "4" : TextBox12.Text = "000000" & "" & TextBox12.Text
                Case "5" : TextBox12.Text = "00000" & "" & TextBox12.Text
                Case "6" : TextBox12.Text = "0000" & "" & TextBox12.Text
                Case "7" : TextBox12.Text = "000" & "" & TextBox12.Text
                Case "8" : TextBox12.Text = "00" & "" & TextBox12.Text
                Case "9" : TextBox12.Text = "0" & "" & TextBox12.Text
                Case "10" : TextBox12.Text = TextBox12.Text
            End Select
            TextBox10.Select()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add(1)
        Dim ui As Integer
        ui = DataGridView1.Rows.Count

        If ui = 1 Then

            DataGridView1.Rows(0).Cells(0).Value = 1
            'DataGridView1.Rows(0).Cells(4).Value = DateTime.Now.ToString("yyyy/MM/dd")
            DataGridView1.Rows(0).Selected = True
            DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(1)
        Else
            DataGridView1.Rows(ui - 1).Cells(0).Value = DataGridView1.Rows(ui - 2).Cells(0).Value + 1
            'DataGridView1.Rows(ui - 1).Cells(4).Value = DateTime.Now.ToString("yyyy/MM/dd")
            DataGridView1.Rows(ui - 1).Selected = True
            DataGridView1.CurrentCell = DataGridView1.Rows(ui - 1).Cells(1)
        End If

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

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
            ComboBox1.SelectedIndex = -1
        End If
    End Sub



    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        abrir()
        enunciado2 = New SqlCommand("select abridora from programa_tenido_detalle where urdido ='" + ComboBox1.Text + "'", conx)
        respuesta2 = enunciado2.ExecuteReader
        While respuesta2.Read
            TextBox1.Text = respuesta2.Item("abridora")
        End While
        respuesta2.Close()

        enunciado4 = New SqlCommand("select tenido from programa_tenido_detalle where urdido ='" + Trim(ComboBox1.Text) + "'", conx)
        respuesta4 = enunciado4.ExecuteReader
        While respuesta4.Read
            TextBox11.Text = respuesta4.Item("tenido")
        End While
        respuesta4.Close()
        enunciado5 = New SqlCommand("select articulo from programa_tenido_detalle where urdido ='" + ComboBox1.Text + "'", conx)
        respuesta5 = enunciado5.ExecuteReader
        While respuesta5.Read
            TextBox2.Text = respuesta5.Item("articulo")
        End While
        respuesta5.Close()

        enunciado6 = New SqlCommand("select titulo from programa_tenido_detalle where urdido ='" + Trim(ComboBox1.Text) + "'", conx)
        respuesta6 = enunciado6.ExecuteReader
        While respuesta6.Read
            TextBox3.Text = respuesta6.Item("titulo")
        End While
        respuesta6.Close()

        enunciado7 = New SqlCommand("select hilo from programa_tenido_detalle where urdido ='" + Trim(ComboBox1.Text) + "'", conx)
        respuesta7 = enunciado7.ExecuteReader
        While respuesta7.Read
            TextBox5.Text = respuesta7.Item("hilo")
        End While
        respuesta7.Close()
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        'ComboBox1.Items.Clear()
        TextBox12.Enabled = True
        DataGridView1.Rows.Clear()
        DataGridView2.DataSource = ""

        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox10.Text = ""
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox5.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        TextBox11.Enabled = True
        TextBox12.Enabled = True
        TextBox10.Enabled = True

        'TextBox14.Enabled = False
        'TextBox13.Enabled = False
        DataGridView1.Enabled = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Abridora_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim func As New fabridora
        Me.TextBox12.Text = func.correlativo_abridora + 1
        Select Case TextBox12.Text.Length
            Case "1" : TextBox12.Text = "000000000" & "" & TextBox12.Text
            Case "2" : TextBox12.Text = "00000000" & "" & TextBox12.Text
            Case "3" : TextBox12.Text = "0000000" & "" & TextBox12.Text
            Case "4" : TextBox12.Text = "000000" & "" & TextBox12.Text
            Case "5" : TextBox12.Text = "00000" & "" & TextBox12.Text
            Case "6" : TextBox12.Text = "0000" & "" & TextBox12.Text
            Case "7" : TextBox12.Text = "000" & "" & TextBox12.Text
            Case "8" : TextBox12.Text = "00" & "" & TextBox12.Text
            Case "9" : TextBox12.Text = "0" & "" & TextBox12.Text
            Case "10" : TextBox12.Text = TextBox12.Text
        End Select
        TextBox10.Select()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub
    Dim ll As New DataTable
    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim jk As New fabridora
            Dim kk As New vabridora
            Dim l As Integer
            Select Case TextBox12.Text.Length
                Case "1" : TextBox12.Text = "000000000" & "" & TextBox12.Text
                Case "2" : TextBox12.Text = "00000000" & "" & TextBox12.Text
                Case "3" : TextBox12.Text = "0000000" & "" & TextBox12.Text
                Case "4" : TextBox12.Text = "000000" & "" & TextBox12.Text
                Case "5" : TextBox12.Text = "00000" & "" & TextBox12.Text
                Case "6" : TextBox12.Text = "0000" & "" & TextBox12.Text
                Case "7" : TextBox12.Text = "000" & "" & TextBox12.Text
                Case "8" : TextBox12.Text = "00" & "" & TextBox12.Text
                Case "9" : TextBox12.Text = "0" & "" & TextBox12.Text
                Case "10" : TextBox12.Text = TextBox12.Text
            End Select
            kk.gid = TextBox12.Text

            ll = jk.buscar_abridora(kk)
            DataGridView1.Rows.Clear()
            If ll.Rows.Count <> 0 Then
                DataGridView2.DataSource = ll
                l = DataGridView2.Rows.Count
                DataGridView1.Rows.Add(l)
                abrir()
                llenar_combo_box2(ComboBox1, Trim(DataGridView2.Rows(0).Cells(12).Value))
                For i = 0 To l - 1
                    TextBox10.Text = Trim(DataGridView2.Rows(0).Cells(12).Value)
                    TextBox2.Text = Trim(DataGridView2.Rows(0).Cells(5).Value)
                    TextBox3.Text = Trim(DataGridView2.Rows(0).Cells(6).Value)
                    TextBox5.Text = Trim(DataGridView2.Rows(0).Cells(7).Value)
                    TextBox7.Text = Trim(DataGridView2.Rows(0).Cells(8).Value)
                    TextBox1.Text = Trim(DataGridView2.Rows(0).Cells(2).Value)
                    TextBox11.Text = Trim(DataGridView2.Rows(0).Cells(3).Value)
                    TextBox8.Text = Trim(DataGridView2.Rows(0).Cells(9).Value)
                    TextBox9.Text = Trim(DataGridView2.Rows(0).Cells(4).Value)
                    TextBox4.Text = Trim(DataGridView2.Rows(0).Cells(10).Value)
                    'TextBox13.Text = DataGridView2.Rows(0).Cells(11).Value

                    DataGridView1.Rows(i).Cells(0).Value = Trim(DataGridView2.Rows(i).Cells(17).Value)
                    DataGridView1.Rows(i).Cells(1).Value = Trim(DataGridView2.Rows(i).Cells(18).Value)
                    DataGridView1.Rows(i).Cells(2).Value = Trim(DataGridView2.Rows(i).Cells(19).Value)
                    DataGridView1.Rows(i).Cells(3).Value = Trim(DataGridView2.Rows(i).Cells(20).Value)
                    DataGridView1.Rows(i).Cells(4).Value = Trim(DataGridView2.Rows(i).Cells(21).Value)
                    DataGridView1.Rows(i).Cells(5).Value = Trim(DataGridView2.Rows(i).Cells(22).Value)
                    DataGridView1.Rows(i).Cells(6).Value = Trim(DataGridView2.Rows(i).Cells(23).Value)
                    DataGridView1.Rows(i).Cells(7).Value = Trim(DataGridView2.Rows(i).Cells(24).Value)
                    DataGridView1.Rows(i).Cells(8).Value = Trim(DataGridView2.Rows(i).Cells(25).Value)
                    DataGridView1.Rows(i).Cells(9).Value = Trim(DataGridView2.Rows(i).Cells(26).Value)
                    DataGridView1.Rows(i).Cells(10).Value = Trim(DataGridView2.Rows(i).Cells(27).Value)
                    ComboBox1.Text = Trim(DataGridView2.Rows(i).Cells(27).Value).ToString
                Next
                TextBox2.Enabled = False
                TextBox3.Enabled = False
                TextBox5.Enabled = False
                TextBox7.Enabled = False
                TextBox8.Enabled = False
                TextBox9.Enabled = False
                TextBox11.Enabled = False
                TextBox12.Enabled = False
                TextBox10.Enabled = False

                'TextBox14.Enabled = False
                'TextBox13.Enabled = False
                DataGridView1.Enabled = False
                PictureBox2.Enabled = True
                PictureBox3.Enabled = True
                PictureBox4.Enabled = True
            Else
                DataGridView1.Rows.Clear()
                DataGridView2.DataSource = ""

                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox5.Text = ""
                TextBox7.Text = ""
                TextBox8.Text = ""
                TextBox9.Text = ""
                TextBox11.Text = ""
                TextBox12.Text = ""
                TextBox10.Text = ""

                'TextBox14.Text = ""
                'TextBox13.Text = ""
            End If

            PictureBox2.Enabled = True
            PictureBox3.Enabled = True
            TextBox1.Enabled = False
        End If
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

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
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
            If DataGridView1.Rows(fila).Cells(5).Value = "" Then
            Else
                Select Case DataGridView1.Rows(fila).Cells(5).Value.Length
                    Case "1" : DataGridView1.Rows(fila).Cells(5).Value = "0000" & "" & DataGridView1.Rows(fila).Cells(5).Value
                    Case "2" : DataGridView1.Rows(fila).Cells(5).Value = "000" & "" & DataGridView1.Rows(fila).Cells(5).Value
                    Case "3" : DataGridView1.Rows(fila).Cells(5).Value = "00" & "" & DataGridView1.Rows(fila).Cells(5).Value
                    Case "4" : DataGridView1.Rows(fila).Cells(5).Value = "0" & "" & DataGridView1.Rows(fila).Cells(5).Value
                    Case "5" : DataGridView1.Rows(fila).Cells(5).Value = DataGridView1.Rows(fila).Cells(5).Value
                End Select
                Select Case DataGridView1.Rows(fila).Cells(5).Value.Length
                    Case "1" : fg.gcodigo = "0000" & "" & DataGridView1.Rows(fila).Cells(5).Value
                    Case "2" : fg.gcodigo = "000" & "" & DataGridView1.Rows(fila).Cells(5).Value
                    Case "3" : fg.gcodigo = "00" & "" & DataGridView1.Rows(fila).Cells(5).Value
                    Case "4" : fg.gcodigo = "0" & "" & DataGridView1.Rows(fila).Cells(5).Value
                    Case "5" : fg.gcodigo = DataGridView1.Rows(fila).Cells(5).Value
                End Select


                DataGridView1.Rows(fila).Cells(6).Value = fg1.buscar_persona(fg)
            End If


        End If
    End Sub

End Class