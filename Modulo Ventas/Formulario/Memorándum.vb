Imports System.Data.SqlClient
Public Class Memorándum
    Public conx As SqlConnection
    Public comando As SqlCommand
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
        DataGridView1.Rows.Add(1)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
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

        If Trim(TextBox3.Text).Length = 0 Then
            MsgBox("FALTA INFORMACION  QUIEN EMITE EL MEMORANDUM")
        Else
            If Trim(TextBox5.Text).Length = 0 Then
                MsgBox("FALTA INFORMACION A QUIEN VA DIRIGID EL MEMORANDUM")
            Else
                If Trim(TextBox2.Text).Length = 0 Then
                    MsgBox("FALTA INFORMACION DEL NUMERO DEL MEMORANDUM")
                Else
                    If DataGridView1.Rows.Count = 0 And Trim(ComboBox1.Text) = "TARDANZA" Then
                        MsgBox("FALTA INFORMACION DEL DETALLE DEL MEMORANDUM")
                    Else

                        Dim agregar As String = "delete from Memorandum_Detalle where CodMem ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "'"
                        Dim agregar1 As String = "delete from Memorandum_cabecera where CodMem ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "'"
                        ELIMINAR(agregar)
                        ELIMINAR(agregar1)
                        Dim cmd20 As New SqlCommand("insert into Memorandum_cabecera (CodMem,DeGMen,CarGMem,AReMem,CarRMem,AsuMem,FecMem,GloMEn,TipMen) values (@CodMem,@DeGMen,@CarGMem,@AReMem,@CarRMem,@AsuMem,@FecMem,@GloMEn,@TipMen)", conx)
                        cmd20.Parameters.AddWithValue("@CodMem", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                        cmd20.Parameters.AddWithValue("@DeGMen", Trim(TextBox3.Text))
                        cmd20.Parameters.AddWithValue("@CarGMem", Trim(TextBox4.Text))
                        cmd20.Parameters.AddWithValue("@AReMem", Trim(TextBox5.Text))
                        cmd20.Parameters.AddWithValue("@CarRMem", Trim(TextBox6.Text))
                        cmd20.Parameters.AddWithValue("@AsuMem", Trim(TextBox7.Text))
                        cmd20.Parameters.AddWithValue("@FecMem", DateTimePicker1.Value)
                        cmd20.Parameters.AddWithValue("@GloMEn", Trim(TextBox8.Text))
                        If Trim(ComboBox1.Text) = "TARDANZA" Then
                            cmd20.Parameters.AddWithValue("@TipMen", "1")
                        Else
                            cmd20.Parameters.AddWithValue("@TipMen", "2")
                        End If

                        cmd20.ExecuteNonQuery()

                        Dim p As Integer
                        p = DataGridView1.Rows.Count

                        If p > 0 Then

                            For i = 0 To p - 1
                                Dim cmd21 As New SqlCommand("insert into Memorandum_Detalle (CodMem,DetMem) values (@CodMem,@DetMem)", conx)
                                cmd21.Parameters.AddWithValue("@CodMem", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                                cmd21.Parameters.AddWithValue("@DetMem", DataGridView1.Rows(i).Cells(0).Value)
                                cmd21.ExecuteNonQuery()
                            Next
                        End If
                        MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
                        Dim respuesta As DialogResult
                        respuesta = MessageBox.Show("DESEA IMPRIMIR EL MEMORANDUM?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        If (respuesta = Windows.Forms.DialogResult.Yes) Then
                            If ComboBox1.Text = "TARDANZA" Then
                                Form_Memorandum.TextBox1.Text = Trim(TextBox1.Text) & Trim(TextBox2.Text)
                                Form_Memorandum.ShowDialog()
                            Else
                                Reporte.TextBox1.Text = Trim(TextBox1.Text) & Trim(TextBox2.Text)
                                Reporte.ShowDialog()
                            End If
                        End If
                        LIMPIAR()
                        corelativo()
                        TextBox8.ReadOnly = True

                    End If
                End If
            End If

        End If


    End Sub
    Sub LIMPIAR()
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        DataGridView1.Rows.Clear()
    End Sub
    Sub bloquear()
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        DataGridView1.Enabled = False
        DateTimePicker1.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
    End Sub
    Sub activar()
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        DataGridView1.Enabled = True
        DateTimePicker1.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
    End Sub
    Sub corelativo()
        abrir()
        Dim sql1023 As String = "select top 1 CodMem from Memorandum_cabecera  order by CodMem desc"
        Dim cmd1023 As New SqlCommand(sql1023, conx)
        Rsr24 = cmd1023.ExecuteReader()
        If Rsr24.Read() = True Then
            TextBox2.Text = Integer.Parse(Mid(Rsr24(0), 5, 6)) + 1

        Else
            TextBox2.Text = 1

        End If
        Rsr24.Close()

        Select Case TextBox2.Text.Length
            Case "1" : TextBox2.Text = "00000" & "" & TextBox2.Text
            Case "1" : TextBox2.Text = "00000" & "" & TextBox2.Text
            Case "2" : TextBox2.Text = "0000" & "" & TextBox2.Text
            Case "3" : TextBox2.Text = "000" & "" & TextBox2.Text
            Case "4" : TextBox2.Text = "00" & "" & TextBox2.Text
            Case "5" : TextBox2.Text = "0" & "" & TextBox2.Text
            Case "6" : TextBox2.Text = TextBox2.Text
        End Select
    End Sub
    Dim Rsr24, Rsr21 As SqlDataReader
    Private Sub Memorándum_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = Year(DateTime.Now)
        TextBox2.Select()
        corelativo()
        bloquear()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.F1 Then
            TRABAJADORES.Label2.Text = 3
            TRABAJADORES.ShowDialog()
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.F1 Then
            TRABAJADORES.Label2.Text = 4
            TRABAJADORES.ShowDialog()
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Not Char.IsDigit(e.KeyChar) And e.KeyChar <> ChrW(Keys.Back) Then
            ' Si no es un número ni la tecla de retroceso, cancelar la entrada
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim estado As Integer
            Select Case TextBox2.Text.Length
                Case "1" : TextBox2.Text = "00000" & TextBox2.Text
                Case "2" : TextBox2.Text = "0000" & TextBox2.Text
                Case "3" : TextBox2.Text = "000" & TextBox2.Text
                Case "4" : TextBox2.Text = "00" & TextBox2.Text
                Case "5" : TextBox2.Text = "0" & TextBox2.Text
                Case "6" : TextBox2.Text = TextBox2.Text

            End Select
            estado = 0
            Dim sql1023 As String = "select * from Memorandum_cabecera where CodMem ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "'"
            Dim cmd1023 As New SqlCommand(sql1023, conx)
            Rsr24 = cmd1023.ExecuteReader()
            If Rsr24.Read() = True Then
                TextBox3.Text = Rsr24(1)
                TextBox4.Text = Rsr24(2)
                TextBox5.Text = Rsr24(3)
                TextBox6.Text = Rsr24(4)
                TextBox7.Text = Rsr24(5)
                DateTimePicker1.Value = Rsr24(6)
                TextBox8.Text = Rsr24(7)
                If Trim(Rsr24(8)) = "1" Then
                    ComboBox1.SelectedIndex = 0
                Else
                    ComboBox1.SelectedIndex = 1
                End If

                estado = 1

            Else
                estado = 2
            End If
            Rsr24.Close()

            If estado = 1 Then
                TextBox2.Enabled = False
                Dim sql1021 As String = "select * from Memorandum_Detalle where CodMem ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "'"
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr21 = cmd1021.ExecuteReader()
                Dim i51 As Integer
                i51 = 0

                While Rsr21.Read()
                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(i51).Cells(0).Value = Rsr21(2)
                    Button3.Enabled = True
                    Button5.Enabled = False
                    Button5.Enabled = False
                    TextBox3.Enabled = False
                    TextBox4.Enabled = False
                    TextBox5.Enabled = False
                    TextBox6.Enabled = False
                    TextBox7.Enabled = False
                    DateTimePicker1.Enabled = False
                    i51 = i51 + 1
                End While
                Rsr21.Close()
                Button3.Enabled = True
            Else
                activar()
                TextBox3.Select()
                TextBox2.Enabled = True
                ComboBox1.Enabled = False
            End If
            ComboBox1.Enabled = True

        End If
        If e.KeyCode = Keys.F1 Then
            Lista_Memorandum.ShowDialog()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ComboBox1.Text = "TARDANZA" Then
            Form_Memorandum.TextBox1.Text = Trim(TextBox1.Text) & Trim(TextBox2.Text)
            Form_Memorandum.ShowDialog()
        Else
            Reporte.TextBox1.Text = Trim(TextBox1.Text) & Trim(TextBox2.Text)
            Reporte.ShowDialog()
        End If

        LIMPIAR()
        corelativo()
        TextBox2.Enabled = True
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        LIMPIAR()
        corelativo()
        TextBox2.Enabled = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "TARDANZA" Then
            TextBox7.Text = "AMONESTACION POR TARDANZAS INJUSTIFICADAS"
            TextBox8.ReadOnly = True
        Else
            TextBox7.Text = "LLAMADO DE ATENCION POR INCUMPLIMIENTO DE FUNCIONES"
            TextBox8.ReadOnly = False
        End If
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button5.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        DateTimePicker1.Enabled = True
        ComboBox1.Enabled = True
        DataGridView1.Enabled = True
        If ComboBox1.Text = "TARDANZA" Then
            TextBox8.ReadOnly = True
        Else
            TextBox8.ReadOnly = False
        End If

    End Sub
End Class