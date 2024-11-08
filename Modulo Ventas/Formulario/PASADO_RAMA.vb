Imports System.Data.SqlClient
Public Class PASADO_RAMA
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim DY As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim jh As New frama
            Dim ll As New vrama
            Dim gg As Integer
            Dim sum As Double
            'Dim sum As Double
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
            ll.gcodigo = TextBox1.Text

            DY = jh.buscar_codigo_rama(ll)

            If DY.Rows.Count > 0 Then
                DataGridView2.DataSource = DY
                gg = DataGridView2.Rows.Count

                DataGridView1.Rows.Add(gg)

                For i = 0 To gg - 1
                    If Trim(DataGridView2.Rows(i).Cells(17).Value) = "" Then
                        DataGridView1.Rows(i).Cells(0).Value = "-"
                    Else
                        DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(17).Value
                    End If
                    If Trim(DataGridView2.Rows(i).Cells(19).Value) = "" Then
                        DataGridView1.Rows(i).Cells(2).Value = "-"
                    Else
                        DataGridView1.Rows(i).Cells(2).Value = DataGridView2.Rows(i).Cells(19).Value
                    End If

                    If DataGridView2.Rows(i).Cells(18).Value = 1 Or DataGridView2.Rows(i).Cells(18).Value = 2 Then
                        DataGridView1.Rows(i).Cells(1).Value = True
                        DataGridView1.Rows(i).ReadOnly = True
                        DataGridView1.Rows(i).Cells(14).ReadOnly = False
                        DataGridView1.Rows(i).Cells(15).ReadOnly = False
                        DataGridView1.Rows(i).Cells(16).ReadOnly = False
                        DataGridView1.Rows(i).Cells(17).ReadOnly = False
                        DataGridView1.Rows(i).Cells(18).ReadOnly = False
                        DataGridView1.Rows(i).Cells(20).ReadOnly = False
                    Else
                        DataGridView1.Rows(i).Cells(1).Value = False
                    End If
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(2).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(4).Value
                    DataGridView1.Rows(i).Cells(6).Value = DataGridView2.Rows(i).Cells(6).Value
                    DataGridView1.Rows(i).Cells(7).Value = DataGridView2.Rows(i).Cells(5).Value
                    DataGridView1.Rows(i).Cells(8).Value = DataGridView2.Rows(i).Cells(7).Value
                    DataGridView1.Rows(i).Cells(9).Value = DataGridView2.Rows(i).Cells(8).Value
                    DataGridView1.Rows(i).Cells(18).Value = DataGridView2.Rows(i).Cells(25).Value
                    DataGridView1.Rows(i).Cells(20).Value = DataGridView2.Rows(i).Cells(26).Value
                    If Trim(DataGridView2.Rows(i).Cells(10).Value) = "" Then
                        DataGridView1.Rows(i).Cells(10).Value = "0.00"
                    Else
                        DataGridView1.Rows(i).Cells(10).Value = DataGridView2.Rows(i).Cells(10).Value
                    End If
                    If Trim(DataGridView2.Rows(i).Cells(11).Value) = "" Then
                        DataGridView1.Rows(i).Cells(11).Value = "0"
                    Else
                        DataGridView1.Rows(i).Cells(11).Value = DataGridView2.Rows(i).Cells(11).Value
                    End If



                    DataGridView1.Rows(i).Cells(12).Value = DataGridView2.Rows(i).Cells(0).Value
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView2.Rows(i).Cells(1).Value
                    DataGridView1.Rows(i).Cells(14).Value = DataGridView2.Rows(i).Cells(20).Value
                    DataGridView1.Rows(i).Cells(15).Value = DataGridView2.Rows(i).Cells(21).Value

                    DataGridView1.Rows(i).Cells(16).Value = DataGridView2.Rows(i).Cells(22).Value


                    DataGridView1.Rows(i).Cells(17).Value = DataGridView2.Rows(i).Cells(23).Value
                    DataGridView1.Rows(i).Cells(19).Value = DataGridView2.Rows(i).Cells(3).Value

                    sum = sum + DataGridView1.Rows(i).Cells(9).Value
                    DataGridView1.Rows(i).Cells(2).ReadOnly = True
                    DataGridView1.Rows(i).Cells(3).ReadOnly = True
                    If DataGridView1.Rows(i).Cells(1).Value = True And DataGridView1.Rows(i).Cells(3).Value = False Then
                        DataGridView1.Rows(i).Cells(3).ReadOnly = False
                    End If
                    If DataGridView2.Rows(i).Cells(18).Value = 2 Then
                        DataGridView1.Rows(i).Cells(3).Value = True
                        DataGridView1.Rows(i).Cells(3).ReadOnly = True
                    Else
                        DataGridView1.Rows(i).Cells(3).Value = False
                    End If
                Next
                TextBox2.Text = sum
                Button1.Enabled = True
                Button2.Enabled = True
            End If
            TextBox1.Enabled = False
            DataGridView2.DataSource = ""
        End If
    End Sub

    'Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
    '    If DataGridView1.Columns(e.ColumnIndex).HeaderText = "S" Then
    '        MsgBox(e.RowIndex - 1)
    '        DataGridView1.Rows(e.RowIndex - 1).Cells(0).Value = DateTime.Now.ToString("hh:mm")

    '    End If
    'End Sub    

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick




    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        DataGridView1.Rows.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox1.Enabled = True
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim jh As New Rpt_Pa_Rama
        jh.TextBox1.Text = TextBox1.Text
        jh.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Observacion_rama.TextBox2.Text = TextBox1.Text
        Observacion_rama.Show()
    End Sub

    Private Sub PASADO_RAMA_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_PaddingChanged(sender As Object, e As EventArgs) Handles TextBox1.PaddingChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If Label4.Text = 1 Then
            MsgBox("EL FORMULARIO YA ESTA ABIERTO")
            Me.Select()
        Else

            Dim pl As Integer
            pl = Label5.Text
            Label4.Text = 1
            UPDATE_HOT.Label4.Text = pl
            UPDATE_HOT.TextBox2.Text = DataGridView1.Rows(pl).Cells(5).Value
            UPDATE_HOT.Show()
        End If


    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub
    Private WithEvents CellTextBox As DataGridViewTextBoxEditingControl
    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        CellTextBox = TryCast(e.Control, DataGridViewTextBoxEditingControl)
    End Sub
    Private Sub CellTextBox_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles CellTextBox.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = Convert.ToChar(Keys.Back) And Not e.KeyChar = Convert.ToChar(Keys.Enter) And Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub
    Dim dy1 As DataTable
    Sub ACTUALIZAR()
        DataGridView2.DataSource = ""
        DataGridView1.Rows.Clear()

        Dim jh As New frama
        Dim ll As New vrama
        Dim gg As Integer
        Dim sum As Double
        'Dim sum As Double
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
        ll.gcodigo = TextBox1.Text

        dy1 = jh.buscar_codigo_rama(ll)

        If DY.Rows.Count > 0 Then
            DataGridView2.DataSource = dy1
            gg = DataGridView2.Rows.Count

            DataGridView1.Rows.Add(gg)

            For i = 0 To gg - 1

                If Trim(DataGridView2.Rows(i).Cells(17).Value) = "" Then
                    DataGridView1.Rows(i).Cells(0).Value = "-"
                Else
                    DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(17).Value
                End If
                If Trim(DataGridView2.Rows(i).Cells(19).Value) = "" Then
                    DataGridView1.Rows(i).Cells(2).Value = "-"
                Else
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView2.Rows(i).Cells(19).Value
                End If
                If DataGridView2.Rows(i).Cells(18).Value = 1 Or DataGridView2.Rows(i).Cells(18).Value = 2 Then
                    DataGridView1.Rows(i).Cells(1).Value = True
                    DataGridView1.Rows(i).ReadOnly = True
                    DataGridView1.Rows(i).Cells(14).ReadOnly = False
                    DataGridView1.Rows(i).Cells(15).ReadOnly = False
                    DataGridView1.Rows(i).Cells(16).ReadOnly = False
                    DataGridView1.Rows(i).Cells(17).ReadOnly = False
                    DataGridView1.Rows(i).Cells(18).ReadOnly = False
                    DataGridView1.Rows(i).Cells(20).ReadOnly = False
                Else
                    DataGridView1.Rows(i).Cells(1).Value = False
                End If
                DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(2).Value
                DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(4).Value
                DataGridView1.Rows(i).Cells(6).Value = DataGridView2.Rows(i).Cells(6).Value
                DataGridView1.Rows(i).Cells(7).Value = DataGridView2.Rows(i).Cells(5).Value
                DataGridView1.Rows(i).Cells(8).Value = DataGridView2.Rows(i).Cells(7).Value
                DataGridView1.Rows(i).Cells(9).Value = DataGridView2.Rows(i).Cells(8).Value
                If Trim(DataGridView2.Rows(i).Cells(10).Value) = "" Then
                    DataGridView1.Rows(i).Cells(10).Value = "0.00"
                Else
                    DataGridView1.Rows(i).Cells(10).Value = DataGridView2.Rows(i).Cells(10).Value
                End If
                If Trim(DataGridView2.Rows(i).Cells(11).Value) = "" Then
                    DataGridView1.Rows(i).Cells(11).Value = "0"
                Else
                    DataGridView1.Rows(i).Cells(11).Value = DataGridView2.Rows(i).Cells(11).Value
                End If



                DataGridView1.Rows(i).Cells(12).Value = DataGridView2.Rows(i).Cells(0).Value
                DataGridView1.Rows(i).Cells(13).Value = DataGridView2.Rows(i).Cells(1).Value
                DataGridView1.Rows(i).Cells(14).Value = DataGridView2.Rows(i).Cells(20).Value
                DataGridView1.Rows(i).Cells(15).Value = DataGridView2.Rows(i).Cells(21).Value
                DataGridView1.Rows(i).Cells(18).Value = DataGridView2.Rows(i).Cells(25).Value
                DataGridView1.Rows(i).Cells(20).Value = DataGridView2.Rows(i).Cells(26).Value
                DataGridView1.Rows(i).Cells(16).Value = DataGridView2.Rows(i).Cells(22).Value


                DataGridView1.Rows(i).Cells(17).Value = DataGridView2.Rows(i).Cells(23).Value
                DataGridView1.Rows(i).Cells(19).Value = DataGridView2.Rows(i).Cells(3).Value

                sum = sum + DataGridView1.Rows(i).Cells(9).Value
                DataGridView1.Rows(i).Cells(2).ReadOnly = True
                DataGridView1.Rows(i).Cells(3).ReadOnly = True

                If DataGridView1.Rows(i).Cells(1).Value = True And DataGridView1.Rows(i).Cells(3).Value = False Then
                    DataGridView1.Rows(i).Cells(3).ReadOnly = False
                End If
                If DataGridView2.Rows(i).Cells(18).Value = 2 Then
                    DataGridView1.Rows(i).Cells(3).Value = True
                    DataGridView1.Rows(i).Cells(3).ReadOnly = True
                Else
                    DataGridView1.Rows(i).Cells(3).Value = False
                End If
            Next
            TextBox2.Text = sum
            Button1.Enabled = True
            Button2.Enabled = True

        End If
        TextBox1.Enabled = False
        DataGridView2.DataSource = ""
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("INGRESE EL NUMERO DE PROGRAMA DE RAMA")
        Else
            ACTUALIZAR()
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim j As Integer

        j = DataGridView1.Rows.Count

        For i = 0 To j - 1
            If DataGridView1.Rows(i).Cells(1).Value = True Or DataGridView1.Rows(i).Cells(3).Value = True Then
                abrir()

                Dim cmd15 As New SqlCommand("update rama set velocidad =@velicidad,temperatura=@temperatura,anchor=@ancho,densidadr=@densidad,sobrealimentacion =@sb,densidad_reposa = @rps WHERE codigo =@CODIGO and id=@id", conx)
                cmd15.Parameters.AddWithValue("@velicidad", Trim(DataGridView1.Rows(i).Cells(16).Value))
                cmd15.Parameters.AddWithValue("@temperatura", Trim(DataGridView1.Rows(i).Cells(17).Value))
                cmd15.Parameters.AddWithValue("@ancho", Trim(DataGridView1.Rows(i).Cells(14).Value))
                cmd15.Parameters.AddWithValue("@densidad", Trim(DataGridView1.Rows(i).Cells(15).Value))
                cmd15.Parameters.AddWithValue("@CODIGO", Trim(DataGridView1.Rows(i).Cells(13).Value))
                cmd15.Parameters.AddWithValue("@id", Trim(DataGridView1.Rows(i).Cells(12).Value))
                cmd15.Parameters.AddWithValue("@sb", Trim(DataGridView1.Rows(i).Cells(18).Value))
                cmd15.Parameters.AddWithValue("@rps", Trim(DataGridView1.Rows(i).Cells(20).Value))
                cmd15.ExecuteNonQuery()
            End If
        Next

        MsgBox("SE ACTUALIZO LA INFORMACION SOLICITADA")
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        DataGridView1.BeginEdit(True)
        Dim J As Integer

        J = DataGridView1.Rows.Count
        For I = 0 To J - 1
            If Me.DataGridView1.CurrentRow.Index.ToString() = I Then
                Label5.Text = I
                DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Yellow
            Else

                DataGridView1.Rows(I).DefaultCellStyle.BackColor = COLOR.White
            End If
        Next
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "S" Then
            If Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value) = "-" Or Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value) <> "" Then
                Dim hh As New vrama
                Dim jh As New frama
                DataGridView1.Rows(e.RowIndex).Cells(0).Value = DateTime.Now.ToString("HH:mm")
                DataGridView1.Rows(e.RowIndex).ReadOnly = True
                hh.gcodigo = TextBox1.Text
                hh.ghora = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                hh.gid = DataGridView1.Rows(e.RowIndex).Cells(12).Value
                hh.ganchor = DataGridView1.Rows(e.RowIndex).Cells(14).Value
                hh.gdensidadr = DataGridView1.Rows(e.RowIndex).Cells(15).Value
                hh.gsalimentacion = DataGridView1.Rows(e.RowIndex).Cells(18).Value
                hh.gdpesado = DataGridView1.Rows(e.RowIndex).Cells(20).Value
                jh.actualizar_RAMA(hh)
                DataGridView1.Rows(e.RowIndex).Cells(3).ReadOnly = False
                DataGridView1.Rows(e.RowIndex).Cells(14).ReadOnly = False
                DataGridView1.Rows(e.RowIndex).Cells(15).ReadOnly = False
                DataGridView1.Rows(e.RowIndex).Cells(16).ReadOnly = False
                DataGridView1.Rows(e.RowIndex).Cells(17).ReadOnly = False
                DataGridView1.Rows(e.RowIndex).Cells(18).ReadOnly = False
                DataGridView1.Rows(e.RowIndex).Cells(20).ReadOnly = False
            End If
        Else
            If DataGridView1.Columns(e.ColumnIndex).HeaderText = "F" Then
                If Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) = "-" Or Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value) <> "" Then
                    DataGridView1.Rows(e.RowIndex).Cells(2).Value = DateTime.Now.ToString("HH:mm")
                    DataGridView1.Rows(e.RowIndex).ReadOnly = True
                    Dim hh As New vrama
                    Dim jh As New frama
                    hh.gcodigo = TextBox1.Text
                    hh.ghora = DataGridView1.Rows(e.RowIndex).Cells(2).Value
                    hh.gid = DataGridView1.Rows(e.RowIndex).Cells(12).Value
                    hh.ganchor = DataGridView1.Rows(e.RowIndex).Cells(14).Value
                    hh.gdensidadr = DataGridView1.Rows(e.RowIndex).Cells(15).Value
                    hh.gsalimentacion = DataGridView1.Rows(e.RowIndex).Cells(18).Value
                    hh.gdpesado = DataGridView1.Rows(e.RowIndex).Cells(20).Value
                    jh.actualizar_RAMA_fin(hh)
                    DataGridView1.Rows(e.RowIndex).Cells(14).ReadOnly = False
                    DataGridView1.Rows(e.RowIndex).Cells(15).ReadOnly = False
                    DataGridView1.Rows(e.RowIndex).Cells(16).ReadOnly = False
                    DataGridView1.Rows(e.RowIndex).Cells(18).ReadOnly = False
                    DataGridView1.Rows(e.RowIndex).Cells(20).ReadOnly = False
                End If

            End If

        End If
    End Sub
End Class