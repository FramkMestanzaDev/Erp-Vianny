Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class TEÑIDO_DE_CUERDAS
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
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim fg As New vpartida
            Dim fg1 As New fpartida
            Select Case TextBox15.Text.Length
                Case "1" : TextBox15.Text = "0000" & "" & TextBox15.Text
                Case "2" : TextBox15.Text = "000" & "" & TextBox15.Text
                Case "3" : TextBox15.Text = "00" & "" & TextBox15.Text
                Case "4" : TextBox15.Text = "0" & "" & TextBox15.Text
                Case "5" : TextBox15.Text = TextBox15.Text
            End Select
            Select Case TextBox15.Text.Length
                Case "1" : fg.gcodigo = "0000" & "" & TextBox15.Text
                Case "2" : fg.gcodigo = "000" & "" & TextBox15.Text
                Case "3" : fg.gcodigo = "00" & "" & TextBox15.Text
                Case "4" : fg.gcodigo = "0" & "" & TextBox15.Text
                Case "5" : fg.gcodigo = TextBox15.Text
            End Select

            TextBox14.Text = fg1.buscar_persona(fg)
        End If

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        DataGridView1.DataSource = ""
        TextBox1.Enabled = True
        TextBox2.Text = ""
        TextBox3.Text = ""

        TextBox5.Text = ""

        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

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
            Button1.Enabled = True
            Button2.Enabled = True
            PictureBox4.Enabled = True
        End If
    End Sub



    Private Sub TEÑIDO_DE_CUERDAS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim func As New ftenidocuerda
        Me.TextBox1.Text = func.correlativo_teñido + 1
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
        PictureBox2.Enabled = False
        PictureBox3.Enabled = False
        PictureBox4.Enabled = False
        Button2.Enabled = False
        Button1.Enabled = False
    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        'abrir()
        'enunciado2 = New SqlCommand("select abridora from programa_tenido_detalle where urdido =" + Trim(ComboBox1.Text), conx)
        'respuesta2 = enunciado2.ExecuteReader
        'While respuesta2.Read
        '    TextBox12.Text = respuesta2.Item("abridora")
        'End While
        'respuesta2.Close()
        'enunciado3 = New SqlCommand("select conera from programa_tenido_detalle where urdido =" + Trim(ComboBox1.Text), conx)
        'respuesta3 = enunciado3.ExecuteReader
        'While respuesta3.Read
        '    TextBox7.Text = respuesta3.Item("conera")
        'End While
        'respuesta3.Close()
        'enunciado4 = New SqlCommand("select tenido from programa_tenido_detalle where urdido =" + Trim(ComboBox1.Text), conx)
        'respuesta4 = enunciado4.ExecuteReader
        'While respuesta4.Read
        '    TextBox11.Text = respuesta4.Item("tenido")
        'End While
        'respuesta4.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        abrir()
        enunciado2 = New SqlCommand("select abridora from programa_tenido_detalle where urdido ='" + ComboBox1.Text + "'", conx)
        respuesta2 = enunciado2.ExecuteReader
        While respuesta2.Read
            TextBox12.Text = respuesta2.Item("abridora")
        End While
        respuesta2.Close()
        enunciado3 = New SqlCommand("select conera from programa_tenido_detalle where urdido ='" + Trim(ComboBox1.Text) + "'", conx)
        respuesta3 = enunciado3.ExecuteReader
        While respuesta3.Read
            TextBox7.Text = respuesta3.Item("conera")
        End While
        respuesta3.Close()
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add(1)
        Dim ui As Integer
        ui = DataGridView1.Rows.Count

        If ui = 1 Then
            DataGridView1.Rows(0).Cells(0).Value = 1

        Else
            DataGridView1.Rows(ui - 1).Cells(0).Value = DataGridView1.Rows(ui - 2).Cells(0).Value + 1


        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim hh As New vtenidocuerda
        Dim hh1 As New vtenidocuerda
        Dim hh2 As New vtenidocuerda
        Dim gg As New ftenidocuerda
        Dim j As Integer
        'eliminar tenidio cuerda

        hh.gid = TextBox1.Text
        gg.eliminar_tenido_cuerda(hh)

        ' insertar cabecera
        hh1.gid = TextBox1.Text
        hh1.gjuego_urdido = ComboBox1.Text
        hh1.gjuego_abridora = TextBox12.Text
        hh1.gjuego_tenido = TextBox11.Text
        hh1.gjuego_conera = TextBox7.Text
        hh1.garticuloAs = TextBox2.Text
        hh1.gtitulo = TextBox3.Text
        hh1.ghilo = TextBox5.Text
        If Trim(TextBox9.Text) = "" Then
            hh1.glongitud = ""
        Else
            hh1.glongitud = TextBox9.Text
        End If

        hh1.glote = TextBox8.Text
        hh1.gcodigo = TextBox15.Text
        hh1.gtrabajador = TextBox14.Text
        hh1.gflag = 1
        hh1.gprograma = TextBox10.Text
        hh1.gobservacion = TextBox13.Text
        hh1.gvelocidad = TextBox4.Text
        hh1.gball = TextBox6.Text
        hh1.gfecha = DateTimePicker1.Value
        'insertar detalle
        If gg.insertar_tenido_cabecera(hh1) = True Then
            j = DataGridView1.Rows.Count

            For i = 0 To j - 1
                hh2.gid = TextBox1.Text
                hh2.gjuego_tenidod = TextBox11.Text

                If Convert.ToString(DataGridView1.Rows(i).Cells(0).Value) = "" Then
                    hh2.gcuerdad = ""
                Else
                    hh2.gcuerdad = DataGridView1.Rows(i).Cells(0).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(1).Value) = "" Then
                    hh2.gballd = ""
                Else
                    hh2.gballd = DataGridView1.Rows(i).Cells(1).Value
                End If

                If Convert.ToString(DataGridView1.Rows(i).Cells(2).Value) = "" Then
                    hh2.gtachod = ""
                Else
                    hh2.gtachod = DataGridView1.Rows(i).Cells(2).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(3).Value) = "" Then
                    hh2.gtitulod = ""
                Else
                    hh2.gtitulod = DataGridView1.Rows(i).Cells(3).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                    hh2.ghilod = ""
                Else
                    hh2.ghilod = DataGridView1.Rows(i).Cells(4).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(5).Value) = "" Then
                    hh2.gloted = ""
                Else
                    hh2.gloted = DataGridView1.Rows(i).Cells(5).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(6).Value) = "" Then
                    hh2.gobservaciond = ""
                Else
                    hh2.gobservaciond = DataGridView1.Rows(i).Cells(6).Value
                End If
                hh2.gflagd = 1
                hh2.gprogramad = TextBox10.Text

                gg.insertar_tenido_detalle(hh2)
            Next
            MsgBox("LA INFORMACION SE REGISTRO CON EXITO")
            TextBox2.Text = ""
            TextBox11.Text = ""
            TextBox10.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            TextBox12.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox15.Text = ""
            TextBox14.Text = ""
            TextBox13.Text = ""
            TextBox4.Text = ""
            TextBox6.Text = ""
            DataGridView1.DataSource = ""
            Dim func As New ftenidocuerda
            Me.TextBox1.Text = func.correlativo_teñido + 1
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

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

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
            TextBox11.Enabled = True
            TextBox12.Enabled = True
            TextBox10.Enabled = True
            TextBox13.Enabled = True
            DataGridView1.Enabled = True
            Button2.Enabled = True
            Button1.Enabled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Dim ll As New DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim jk As New ftenidocuerda
            Dim kk As New vtenidocuerda
            Dim l As Integer
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
            kk.gid = TextBox1.Text

            ll = jk.buscar_tenidoooo(kk)

            If ll.Rows.Count <> 0 Then
                DataGridView2.DataSource = ll
                l = DataGridView2.Rows.Count
                DataGridView1.Rows.Add(l)

                For i = 0 To l - 1
                    TextBox10.Text = DataGridView2.Rows(0).Cells(13).Value
                    TextBox2.Text = DataGridView2.Rows(0).Cells(5).Value
                    TextBox11.Text = DataGridView2.Rows(0).Cells(3).Value
                    TextBox3.Text = DataGridView2.Rows(0).Cells(6).Value
                    TextBox5.Text = DataGridView2.Rows(0).Cells(7).Value
                    TextBox12.Text = DataGridView2.Rows(0).Cells(2).Value
                    TextBox7.Text = DataGridView2.Rows(0).Cells(4).Value
                    TextBox8.Text = DataGridView2.Rows(0).Cells(9).Value
                    TextBox9.Text = DataGridView2.Rows(0).Cells(8).Value
                    TextBox15.Text = DataGridView2.Rows(0).Cells(10).Value
                    TextBox14.Text = DataGridView2.Rows(0).Cells(11).Value
                    TextBox13.Text = DataGridView2.Rows(0).Cells(14).Value
                    TextBox4.Text = DataGridView2.Rows(0).Cells(15).Value
                    TextBox6.Text = DataGridView2.Rows(0).Cells(16).Value
                    DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(21).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(i).Cells(22).Value
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView2.Rows(i).Cells(23).Value
                    DataGridView1.Rows(i).Cells(3).Value = DataGridView2.Rows(i).Cells(24).Value
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(25).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(26).Value
                    DataGridView1.Rows(i).Cells(6).Value = DataGridView2.Rows(i).Cells(29).Value
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
                TextBox15.Enabled = False
                TextBox14.Enabled = False
                TextBox13.Enabled = False
                DataGridView1.Enabled = False
                PictureBox2.Enabled = True
                PictureBox3.Enabled = True
                PictureBox4.Enabled = True
                abrir()
                llenar_combo_box2(ComboBox1, TextBox10.Text)
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
                TextBox15.Text = ""
                TextBox14.Text = ""
                TextBox13.Text = ""
                TextBox4.Text = ""
                TextBox6.Text = ""
            End If

            PictureBox2.Enabled = True
            PictureBox3.Enabled = True
            TextBox1.Enabled = False
        End If
    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        If TypeOf e.Control Is TextBox Then

            DirectCast(e.Control, TextBox).CharacterCasing = CharacterCasing.Upper

        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick

    End Sub



    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
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

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'Dim fila As Integer
        'fila = e.RowIndex
        'MsgBox(fila)
        'DataGridView1.Rows(fila).DefaultCellStyle.BackColor = Color.Yellow



    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click

    End Sub
End Class