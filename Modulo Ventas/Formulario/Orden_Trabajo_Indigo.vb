Public Class Orden_Trabajo_Indigo
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim k As Integer
        k = TextBox2.Text
        DataGridView1.Rows.Add(k)
        TextBox4.Text = TextBox2.Text
        TextBox2.Text = ""
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Orden_Trabajo_Indigo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim jj As New fprogramatenido
        Dim l As String
        Label5.Visible = False
        l = jj.correlativo_programa() + 1

        Select Case l.Length
            Case "1" : TextBox1.Text = "000000000" & "" & l
            Case "2" : TextBox1.Text = "00000000" & "" & l
            Case "3" : TextBox1.Text = "0000000" & "" & l
            Case "4" : TextBox1.Text = "000000" & "" & l
            Case "5" : TextBox1.Text = "00000" & "" & l
            Case "6" : TextBox1.Text = "0000" & "" & l
            Case "7" : TextBox1.Text = "000" & "" & l
            Case "8" : TextBox1.Text = "00" & "" & l
            Case "9" : TextBox1.Text = "0" & "" & l
            Case "10" : TextBox1.Text = l
        End Select
        PictureBox2.Enabled = False
        PictureBox3.Enabled = False
        Button3.Enabled = False
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Dim dt As New DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then

            'DataGridView1.Rows.Clear()
            Dim jj As New vprohgrama
            Dim ff As New fprogramatenido
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
            jj.gprograma = TextBox1.Text
            dt = ff.buscar_programa(jj)

            If dt.Rows.Count <> 0 Then
                DataGridView2.DataSource = dt
                l = DataGridView2.Rows.Count
                DataGridView1.Rows.Add(l)
                For i = 0 To l - 1
                    DateTimePicker1.Value = DataGridView2.Rows(0).Cells(2).Value
                    TextBox3.Text = DataGridView2.Rows(0).Cells(3).Value
                    DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(6).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(i).Cells(7).Value
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView2.Rows(i).Cells(8).Value
                    DataGridView1.Rows(i).Cells(3).Value = DataGridView2.Rows(i).Cells(9).Value
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(12).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(13).Value
                    DataGridView1.Rows(i).Cells(6).Value = DataGridView2.Rows(i).Cells(14).Value
                Next
                If DataGridView2.Rows(0).Cells(4).Value = 1 Then
                    Label5.Text = "ACTIVO"
                    Label5.Visible = True
                Else
                    Label5.Visible = True
                End If
                Button3.Enabled = False
                PictureBox2.Enabled = True
            Else
                DataGridView1.Rows.Clear()
                DataGridView1.Enabled = True
                TextBox3.Enabled = True
                TextBox2.Enabled = True
                Button1.Enabled = True
                Label5.Text = "ACTIVO"
                Button3.Enabled = True
            End If
            TextBox2.Select()
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If Label5.Text = "ANULADO" Then
            MsgBox("NO SE PUEDE EDITAR PORQUE EL PROGRAMA ESTA ANULADO")
        Else
            DataGridView1.Enabled = True
            TextBox3.Enabled = True
            TextBox2.Enabled = True
            Button1.Enabled = True
            Button3.Enabled = True
        End If

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim kk As New vprohgrama
        Dim kk1 As New vprohgrama
        Dim kk2 As New vprohgrama
        Dim jj As New fprogramatenido
        If TextBox3.Text = "" Then
            MsgBox("FALTA INGRESAR CANTIDAD A PRODUCIR")
        Else
            kk2.gprograma = TextBox1.Text
            jj.eliminar_programa(kk2)
            kk.gprograma = TextBox1.Text
            kk.gcargadas = TextBox4.Text
            kk.gfecha = DateTimePicker1.Value
            kk.gkg = TextBox3.Text
            kk.gflag = 1

            'jj.insertar_programa(kk)
            Dim h As Integer
            h = DataGridView1.Rows.Count
            If jj.insertar_programa(kk) = True Then

                For i = 0 To h - 1
                    kk1.gabridorad = DataGridView1.Rows(i).Cells(0).Value
                    kk1.gurdidod = DataGridView1.Rows(i).Cells(1).Value
                    kk1.gconerad = DataGridView1.Rows(i).Cells(2).Value
                    kk1.gtenidod = DataGridView1.Rows(i).Cells(3).Value
                    kk1.gprogramad = TextBox1.Text
                    kk1.gflagd = 1
                    kk1.garticulo = DataGridView1.Rows(i).Cells(4).Value
                    kk1.gtitulo = DataGridView1.Rows(i).Cells(5).Value
                    kk1.ghilo = DataGridView1.Rows(i).Cells(6).Value
                    jj.insertar_programa_detalle(kk1)
                Next
                DataGridView1.Rows.Clear()
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                MsgBox("SE REGISTRARON LOS DATOS CORRECTAMENTE")

                Dim l As String

                l = jj.correlativo_programa() + 1

                Select Case l.Length
                    Case "1" : TextBox1.Text = "000000000" & "" & l
                    Case "2" : TextBox1.Text = "00000000" & "" & l
                    Case "3" : TextBox1.Text = "0000000" & "" & l
                    Case "4" : TextBox1.Text = "000000" & "" & l
                    Case "5" : TextBox1.Text = "00000" & "" & l
                    Case "6" : TextBox1.Text = "0000" & "" & l
                    Case "7" : TextBox1.Text = "000" & "" & l
                    Case "8" : TextBox1.Text = "00" & "" & l
                    Case "9" : TextBox1.Text = "0" & "" & l
                    Case "10" : TextBox1.Text = l
                End Select
            End If
        End If

    End Sub



    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        DataGridView1.BeginEdit(True)
    End Sub



    Private Sub DataGridView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DataGridView1.KeyPress
        e.Handled = True
        SendKeys.Send(e.KeyChar.ToString().ToUpper())
    End Sub



    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        If TypeOf e.Control Is TextBox Then

            DirectCast(e.Control, TextBox).CharacterCasing = CharacterCasing.Upper

        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Dim jj As New fprogramatenido

        Dim respuesta As DialogResult
        Dim ml As New vprohgrama
        Dim mk As New fprogramatenido
        ml.gprograma = TextBox1.Text

        respuesta = MessageBox.Show("QUIERES ANULAR PROGRAMA DE TEÑIDO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then


            mk.anular_program_tenido(ml)
            TextBox2.Text = ""
            TextBox3.Text = ""
            DataGridView1.Rows.Clear()
            Dim l As String

            l = jj.correlativo_programa() + 1

            Select Case l.Length
                Case "1" : TextBox1.Text = "000000000" & "" & l
                Case "2" : TextBox1.Text = "00000000" & "" & l
                Case "3" : TextBox1.Text = "0000000" & "" & l
                Case "4" : TextBox1.Text = "000000" & "" & l
                Case "5" : TextBox1.Text = "00000" & "" & l
                Case "6" : TextBox1.Text = "0000" & "" & l
                Case "7" : TextBox1.Text = "000" & "" & l
                Case "8" : TextBox1.Text = "00" & "" & l
                Case "9" : TextBox1.Text = "0" & "" & l
                Case "10" : TextBox1.Text = l
            End Select
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class