Imports System.ComponentModel
Imports System.Data.SqlClient
Public Class Calidad_tela_cruda
    Dim dt, dt2 As New DataTable
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Det_Rollo.Label1.Text = "AUDITOR"
        Det_Rollo.Label2.Text = 3
        Det_Rollo.Show()

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then

            Dim l, l1 As New vtejeduria
            Dim kl As New ftejeduria
            Dim kk As New vpesadorollo
            Dim kk1 As New vpesadorollo
            Dim dg As String
            dg = Mid(Year(DateTimePicker1.Value), 3, 2)
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
                DataGridView2.DataSource = dt

                'If DataGridView2.Rows(0).Cells(4).Value = "PESADO  " Then
                TextBox7.Text = DataGridView2.Rows(0).Cells(8).Value
                    TextBox8.Text = DataGridView2.Rows(0).Cells(44).Value
                    TextBox9.Text = Replace(DataGridView2.Rows(0).Cells(7).Value, "_", " ")
                    TextBox10.Text = DataGridView2.Rows(0).Cells(25).Value
                    TextBox11.Text = DataGridView2.Rows(0).Cells(39).Value
                    TextBox12.Text = DataGridView2.Rows(0).Cells(28).Value
                    TextBox13.Text = DataGridView2.Rows(0).Cells(5).Value
                    TextBox14.Text = DataGridView2.Rows(0).Cells(20).Value
                    TextBox15.Text = DataGridView2.Rows(0).Cells(45).Value
                    TextBox16.Text = DataGridView2.Rows(0).Cells(9).Value
                    TextBox17.Text = DataGridView2.Rows(0).Cells(14).Value
                    TextBox18.Text = DataGridView2.Rows(0).Cells(21).Value
                    TextBox19.Text = DataGridView2.Rows(0).Cells(6).Value
                    TextBox20.Text = DataGridView2.Rows(0).Cells(22).Value
                    TextBox21.Text = DataGridView2.Rows(0).Cells(42).Value
                    TextBox22.Text = DataGridView2.Rows(0).Cells(40).Value
                    TextBox2.Text = DataGridView2.Rows(0).Cells(9).Value
                    If Trim(DataGridView2.Rows(0).Cells(15).Value) = 1 Then
                        TextBox3.Text = "DIURNO"
                    Else
                        TextBox3.Text = "NOCTURNO"
                    End If
                    TextBox5.Text = DataGridView2.Rows(0).Cells(20).Value
                    TextBox6.Text = DataGridView2.Rows(0).Cells(46).Value
                    TextBox23.Text = DataGridView2.Rows(0).Cells(63).Value
                    TextBox29.Text = DataGridView2.Rows(0).Cells(54).Value

                Select Case Trim(DataGridView2.Rows(0).Cells(26).Value)
                    Case "A" : ComboBox1.SelectedIndex = 0
                    Case "B" : ComboBox1.SelectedIndex = 1
                    Case "C" : ComboBox1.SelectedIndex = 2
                    Case "" : ComboBox1.SelectedIndex = 0
                End Select
                If Trim(DataGridView2.Rows(0).Cells(67).Value).Length = 0 Then
                        TextBox25.Text = ""
                    Else
                        kk1.gtejedor = DataGridView2.Rows(0).Cells(67).Value
                        TextBox25.Text = kl.MOSTRAR_TEJEDOR(kk1)
                    End If


                l1.gccia = Label28.Text
                l1.grolloini = TextBox1.Text
                    l1.grollofin = TextBox1.Text
                    dt2 = kl.reporte_defectos_rollo(l1)
                    If dt2.Rows.Count <> 0 Then
                        DataGridView1.DataSource = dt2
                        DataGridView1.Columns(0).Visible = False
                        DataGridView1.Columns(1).Visible = False
                        DataGridView1.Columns(2).Visible = False
                        DataGridView1.Columns(3).Visible = False
                        DataGridView1.Columns(4).Width = 250
                        DataGridView1.Columns(5).Width = 150
                        DataGridView1.Columns(6).Width = 80
                        DataGridView1.Columns(7).Width = 150
                    DataGridView1.Columns(8).Width = 150

                End If

                    kk.gtejedor = DataGridView2.Rows(0).Cells(30).Value
                    TextBox4.Text = kl.MOSTRAR_TEJEDOR(kk)

                    ComboBox1.Enabled = True
                    TextBox23.Enabled = True
                    Button2.Enabled = True
                    DataGridView1.ReadOnly = False
                    Dim j As Integer
                    j = DataGridView1.Rows.Count
                    Dim suma, suma1, suma2 As Double
                    For i = 0 To j - 1
                        suma = suma + DataGridView1.Rows(i).Cells(6).Value
                        suma1 = suma1 + DataGridView1.Rows(i).Cells(7).Value
                        suma2 = suma2 + DataGridView1.Rows(i).Cells(8).Value
                    Next
                    TextBox27.Text = suma
                    TextBox26.Text = suma1
                    TextBox24.Text = suma2

                    If Trim(TextBox6.Text) = "ACTIVO" Then
                        MsgBox("LA PARTIDA YA SE VALIDO PARA PASAR POR RAMA NO PUEDE EDITAR LA INFORMACION DEL ROLLO")
                        ComboBox1.Enabled = False
                        Button2.Enabled = False
                        TextBox23.Enabled = False
                        TextBox29.Enabled = False
                        Button1.Enabled = True
                        DataGridView1.Enabled = False
                    Else
                        ComboBox1.Enabled = True
                        Button2.Enabled = True
                        TextBox23.Enabled = True
                        TextBox29.Enabled = True
                        Button1.Enabled = True
                        DataGridView1.Enabled = True
                    End If
                'Else
                '    MsgBox("EL ROLLO NO ESTA PESADO, NO SE PUEDE AUDITAR")
                'End If

            End If
        End If


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Me.ValidateChildren() And TextBox25.Text = String.Empty Or TextBox1.Text = String.Empty Then
            MsgBox("FALTA INGRESAR CAMPOS OBLIGATORIOS")
        Else
            Dim kl As New ftejeduria
            Dim kj As New vtejeduria
            Dim kj1 As New vtejeduria
            abrir()

            Dim cmd As New SqlCommand("UPDATE custom_vianny.dbo.marg0001 SET flag_3r = '3',obscalidad_3r = @obscalidad_3r, calidad_3r = @calidad_3r,  audicrudo_3r =@audicrudo_3r,frevcrudo_3r =@fecha,kgsafec_3a=@kgsafec_3a where  rollo_3r =@rollo and ccia_3r =@ccia", conx)
            cmd.Parameters.AddWithValue("@fecha", DateTimePicker1.Value)
            cmd.Parameters.AddWithValue("@obscalidad_3r", Trim(TextBox23.Text))
            cmd.Parameters.AddWithValue("@calidad_3r", Trim(ComboBox1.Text))
            cmd.Parameters.AddWithValue("@audicrudo_3r", Trim(TextBox28.Text))
            cmd.Parameters.AddWithValue("@rollo", TextBox1.Text)
            cmd.Parameters.AddWithValue("@ccia", Label28.Text)
            If TextBox29.Text = "" Then
                cmd.Parameters.AddWithValue("@kgsafec_3a", 0)
            Else
                cmd.Parameters.AddWithValue("@kgsafec_3a", TextBox29.Text)
            End If


            cmd.ExecuteNonQuery()
            kj1.gtdr_ccia = Label28.Text
            kj1.gtsdr_rollo = TextBox1.Text
            kj1.gtdr_defecto = ""
            kl.eliminar_defectos(kj1)

            Dim t As Integer
            t = DataGridView1.Rows.Count


            For i = 0 To t - 1
                kj.gtdr_ccia = Label28.Text
                kj.gtsdr_rollo = DataGridView1.Rows(i).Cells(1).Value
                kj.gtdr_defecto = DataGridView1.Rows(i).Cells(2).Value
                kj.gtdr_obs = ""
                kj.gtdr_cantidad = DataGridView1.Rows(i).Cells(6).Value
                kj.gtdr_mtrafectados = DataGridView1.Rows(i).Cells(7).Value
                kj.gtdr_kgafectados = DataGridView1.Rows(i).Cells(8).Value
                kl.actualizar_defectos(kj)
            Next
            MessageBox.Show("DATOS INGRESADOS CORRECTAMENTE")
            eliminar()
        End If
    End Sub





    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Sub eliminar()
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox22.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox23.Text = ""
        TextBox25.Text = ""
        TextBox28.Text = ""
        TextBox4.Text = ""
        TextBox24.Text = ""
        TextBox26.Text = ""
        TextBox27.Text = ""
        TextBox28.Text = ""
        TextBox29.Text = ""
        DataGridView1.DataSource = ""
        DataGridView2.DataSource = ""
        TextBox1.Text = ""
    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        eliminar()
    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click
        calculo()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        calculo()
    End Sub

    Private Sub Calidad_tela_cruda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        calculo()
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged

    End Sub

    Sub calculo()
        Dim j As Integer
        j = DataGridView1.Rows.Count
        Dim suma, suma1, suma2 As Double
        For i = 0 To j - 1
            suma = suma + DataGridView1.Rows(i).Cells(6).Value
            suma1 = suma1 + DataGridView1.Rows(i).Cells(7).Value
            suma2 = suma2 + DataGridView1.Rows(i).Cells(8).Value
        Next
        TextBox27.Text = suma
        TextBox26.Text = suma1
        TextBox24.Text = suma2
    End Sub

    Private Sub TextBox29_TextChanged(sender As Object, e As EventArgs) Handles TextBox29.TextChanged

    End Sub

    Private Sub TextBox25_Validating(sender As Object, e As CancelEventArgs) Handles TextBox25.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.errores.SetError(sender, "")
        Else
            Me.errores.SetError(sender, "FALTA INGRESAR EL AUDITOR")
        End If
    End Sub

    Private Sub TextBox29_Validating(sender As Object, e As CancelEventArgs) Handles TextBox29.Validating
        If Val(TextBox29.Text) = 0 Then
            Me.errores.SetError(sender, "")
        Else
            Me.errores.SetError(sender, "FALTA INGRESAR LOS KG DE RETASOS")
        End If
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
    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.errores.SetError(sender, "")
        Else
            Me.errores.SetError(sender, "FALTA INGRESAR EL ROLLO")
        End If
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

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        NumConFrac(TextBox1, e)
    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        calculo()
    End Sub

    Private Sub TextBox29_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox29.KeyPress
        NumConFrac1(TextBox29, e)
    End Sub
    Public Sub NumConFrac1(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
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
End Class