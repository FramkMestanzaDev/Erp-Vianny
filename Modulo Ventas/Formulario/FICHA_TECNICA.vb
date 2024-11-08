Imports System.ComponentModel
Imports System.Data.SqlClient
Public Class FICHA_TECNICA
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
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
        DataGridView1.Rows.Add()
        Dim I18, VAL As Integer
        I18 = DataGridView1.Rows.Count
        For i1 = 0 To I18 - 1
            VAL = i1 + 1
            DataGridView1.Rows(i1).Cells(0).Value = VAL

        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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
            Dim p As Integer
            p = DataGridView1.Rows.Count
            Dim fg As Double
            For i = 0 To p - 1
                fg = fg + DataGridView1.Rows(i).Cells(4).Value
            Next
            TextBox28.Text = fg.ToString("N2")
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DataGridView2.Rows.Add()

    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView2.Rows.Remove(DataGridView2.CurrentRow)
            I18 = DataGridView2.Rows.Count
            For i1 = 0 To I18 - 1
                VAL = i1 + 1
                DataGridView2.Rows(i1).Cells(0).Value = VAL

            Next
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Dim Rsr2, Rsr20, Rsr21, Rsr22 As SqlDataReader
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            DET_FICHA_TECNICA.Close()
            DET_FICHA_TECNICA.Label2.Text = Label39.Text
            DET_FICHA_TECNICA.Label3.Text = 1
            DET_FICHA_TECNICA.ShowDialog()
        Else
            If e.KeyCode = Keys.Enter Then
                abrir()
                DataGridView1.Rows.Clear()
                DataGridView2.Rows.Clear()

                Dim sql1022 As String = "select COUNT(articulo_17F) as valor from custom_vianny.dbo.Cag1700_Ficha where articulo_17F ='" + Trim(TextBox1.Text) + "' and ccia_17f ='01'"
                Dim cmd1022 As New SqlCommand(sql1022, conx)
                Rsr22 = cmd1022.ExecuteReader()
                Rsr22.Read()
                If Rsr22(0) = 0 Then
                    Rsr22.Close()
                    TabControl1.Enabled = True
                    GroupBox2.Enabled = True
                    GroupBox3.Enabled = True
                    GroupBox4.Enabled = True
                    GroupBox5.Enabled = True
                    TextBox3.Enabled = True
                    TextBox4.Enabled = True
                    TextBox5.Enabled = True
                    TextBox6.Enabled = True
                    TextBox7.Enabled = True
                    TextBox8.Enabled = True
                    TextBox29.Enabled = True
                    Button3.Enabled = True
                    Button4.Enabled = True
                Else
                    Button4.Enabled = True
                    Rsr22.Close()
                    Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAllFichaTextilTela '01','" + Trim(TextBox1.Text) + "',1"
                    Dim cmd1021 As New SqlCommand(sql1021, conx)
                    Rsr21 = cmd1021.ExecuteReader()
                    If Rsr21.Read() Then

                        DateTimePicker1.Value = Rsr21(2)
                        TextBox3.Text = Trim(Rsr21(4))
                        TextBox4.Text = Rsr21(5)
                        TextBox5.Text = (Rsr21(6))
                        TextBox6.Text = (Rsr21(7))
                        TextBox7.Text = (Rsr21(8))
                        TextBox8.Text = (Rsr21(9))

                        TextBox9.Text = Trim(Rsr21(10))
                        TextBox10.Text = Trim(Rsr21(11))
                        If Rsr21(12) = "A" Then
                            ComboBox1.SelectedIndex = 0
                        Else
                            ComboBox1.SelectedIndex = 1
                        End If
                        TextBox11.Text = Rsr21(13)
                        TextBox12.Text = Rsr21(14)
                        If Rsr21(15) = "A" Then
                            ComboBox2.SelectedIndex = 0
                        Else
                            ComboBox2.SelectedIndex = 1
                        End If
                        TextBox13.Text = Rsr21(16)
                        TextBox14.Text = Rsr21(17)
                        If Rsr21(18) = "A" Then
                            ComboBox3.SelectedIndex = 0
                        Else
                            ComboBox3.SelectedIndex = 1
                        End If
                        TextBox15.Text = Rsr21(20)
                        TextBox16.Text = Rsr21(23)
                        TextBox17.Text = Rsr21(26)

                        TextBox18.Text = Trim(Rsr21(28))
                        TextBox19.Text = Trim(Rsr21(29))
                        TextBox20.Text = Trim(Rsr21(30))
                        TextBox21.Text = Trim(Rsr21(31))
                        TextBox22.Text = Trim(Rsr21(32))
                        TextBox23.Text = Trim(Rsr21(33))
                        TextBox24.Text = Trim(Rsr21(34))
                        TextBox25.Text = Trim(Rsr21(35))
                        TextBox26.Text = Trim(Rsr21(36))
                        TextBox27.Text = Trim(Rsr21(37))
                        TextBox29.Text = Trim(Rsr21(38))
                    End If
                    Rsr21.Close()

                    Dim sql10210 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAllFichaTextilTelaDetalle '01','" + Trim(TextBox1.Text) + "'"
                    Dim cmd10210 As New SqlCommand(sql10210, conx)
                    Rsr20 = cmd10210.ExecuteReader()
                    Dim i51 As Integer
                    Dim SUM As Double
                    i51 = 0
                    While Rsr20.Read()
                        DataGridView1.Rows.Add()
                        DataGridView1.Rows(i51).Cells(0).Value = Trim(Rsr20(0))
                        DataGridView1.Rows(i51).Cells(2).Value = Trim(Rsr20(2))
                        DataGridView1.Rows(i51).Cells(3).Value = Trim(Rsr20(13))
                        DataGridView1.Rows(i51).Cells(4).Value = Trim(Rsr20(12))
                        i51 = i51 + 1
                        SUM = SUM + Rsr20(12)
                    End While
                    TextBox28.Text = SUM.ToString("N2")
                    Rsr20.Close()

                    Dim sql102101 As String = "EXEC custom_vianny.dbo.CaeSoft_GetallTelaOP '01','" + Trim(TextBox1.Text) + "',1,NULL"
                    Dim cmd102101 As New SqlCommand(sql102101, conx)
                    Rsr2 = cmd102101.ExecuteReader()
                    Dim i5 As Integer
                    i5 = 0
                    While Rsr2.Read()
                        DataGridView2.Rows.Add()
                        DataGridView2.Rows(i5).Cells(1).Value = Trim(Rsr2(1))
                        DataGridView2.Rows(i5).Cells(2).Value = Trim(Rsr2(2))
                        DataGridView2.Rows(i5).Cells(3).Value = Trim(Rsr2(4))
                        DataGridView2.Rows(i5).Cells(4).Value = Trim(Rsr2(3))
                        i5 = i5 + 1

                    End While
                    Rsr2.Close()

                End If
            End If
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
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TabControl1.Enabled = True
        GroupBox2.Enabled = True
        GroupBox3.Enabled = True
        GroupBox4.Enabled = True
        GroupBox5.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        TextBox29.Enabled = True
        Button3.Enabled = True
        TextBox1.Enabled = False
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = "111"
            Clientes.Show()
        End If
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 1 Then
            DET_FICHA_TECNICA.Close()
            DET_FICHA_TECNICA.Label3.Text = 2
            DET_FICHA_TECNICA.Label4.Text = e.RowIndex

            DET_FICHA_TECNICA.Label5.Visible = True
            DET_FICHA_TECNICA.TextBox2.Visible = True
            DET_FICHA_TECNICA.ShowDialog()
        End If
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "%" Then
            Dim p As Integer
            p = DataGridView1.Rows.Count
            Dim fg As Double
            For i = 0 To p - 1
                fg = fg + DataGridView1.Rows(i).Cells(4).Value
            Next
            TextBox28.Text = fg.ToString("N2")
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.ColumnIndex = 0 Then
            DET_FICHA_TECNICA.Close()
            DET_FICHA_TECNICA.Label3.Text = 3
            DET_FICHA_TECNICA.Label4.Text = e.RowIndex
            DET_FICHA_TECNICA.ShowDialog()
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        Try
            NumConFrac(TextBox9, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        Try
            NumConFrac(TextBox11, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox13_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox13.KeyPress
        Try
            NumConFrac(TextBox13, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        Try
            NumConFrac(TextBox5, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        Try
            NumConFrac(TextBox6, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try

    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        Try
            NumConFrac(TextBox7, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try

    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        Try
            NumConFrac(TextBox8, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try

    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        Try
            NumConFrac1(TextBox10, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress
        Try
            NumConFrac1(TextBox12, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox14_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox14.KeyPress
        Try
            NumConFrac1(TextBox14, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox15_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox15.KeyPress
        Try
            NumConFrac1(TextBox15, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox16_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox16.KeyPress
        Try
            NumConFrac1(TextBox16, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox17_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox17.KeyPress
        Try
            NumConFrac1(TextBox17, e)
        Catch ex As Exception
            MsgBox("FALATA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub TextBox15_LostFocus(sender As Object, e As EventArgs) Handles TextBox15.LostFocus
        Dim NumAuxiliar As Double
        NumAuxiliar = TextBox15.Text
        TextBox15.Text = FormatNumber(NumAuxiliar, 2)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub

    Private Sub TextBox16_LostFocus(sender As Object, e As EventArgs) Handles TextBox16.LostFocus
        Dim NumAuxiliar As Double
        NumAuxiliar = TextBox16.Text
        TextBox16.Text = FormatNumber(NumAuxiliar, 2)
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button24_Click_1(sender As Object, e As EventArgs) Handles Button24.Click
        OpenFileDialog1.ShowDialog()
        TextBox27.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        Process.Start(Trim(TextBox27.Text))
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
    Private Function Validar_Textobox() As Boolean
        Dim metros As Integer
        Dim kilos As Integer


        If Convert.ToDouble(TextBox11.Text) = 0 Then
            MessageBox.Show("La Densidad B/W tiena que ser Mayor a 0", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If

        If Convert.ToDouble(TextBox12.Text) = 0 Then
            MessageBox.Show("El Ancho B/W tiena que ser Mayor a 0", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If Trim(TextBox14.Text).Length = 0 Then
            MessageBox.Show("Falta Ingresar el Nombre Comercial", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If Trim(TextBox1.Text).Length = 0 Then
            MessageBox.Show("Falta Ingresar el Codigo de la Tela", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("Se Tiene que ingresar Informacion de la Composicion de la tela", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If

        Return True
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        abrir()
        If Validar_Textobox() = True Then
            Dim at1, at3, at2 As String
            Dim cmd15 As New SqlCommand("EXEC custom_vianny.dbo.CaeSoft_FichaTextilTelaInsertUpdated '01', @codigo,@fecha,@ruc, @ncomercial,@GALGA,@DIAMETRO,@SISTEMAS,@AGUJAS,@densidad1,@ancho1,@at1,
		@densidad2,@ancho2,@at2,@densidad3,@ancho3,@at3,0.00,@lm1, 0.00 , 0.00 ,@lm2, 0.00 , 0.00 ,@lm3,0.00 ,@ruta1 , 
		@ruta2 ,@ruta3,@ruta4,@ruta5,@ruta6,@ruta7 ,@ruta8,@ruta9,@ruta10,@observacion, '01', '1', '1'", conx)

            If ComboBox1.Text = "ABIERTO" Then
                cmd15.Parameters.AddWithValue("@at1", "A")
            Else
                If ComboBox1.Text = "TUBULAR" Then
                    cmd15.Parameters.AddWithValue("@at1", "T")
                Else
                    cmd15.Parameters.AddWithValue("@at1", "")
                End If

            End If
            If ComboBox2.Text = "ABIERTO" Then
                cmd15.Parameters.AddWithValue("@at2", "A")
            Else
                If ComboBox2.Text = "TUBULAR" Then
                    cmd15.Parameters.AddWithValue("@at2", "T")
                Else
                    cmd15.Parameters.AddWithValue("@at2", "")
                End If

            End If
            If ComboBox3.Text = "ABIERTO" Then
                cmd15.Parameters.AddWithValue("@at3", "A")
            Else
                If ComboBox3.Text = "TUBULAR" Then
                    cmd15.Parameters.AddWithValue("@at3", "T")
                Else
                    cmd15.Parameters.AddWithValue("@at3", "")
                End If

            End If
            cmd15.Parameters.AddWithValue("@codigo", Trim(TextBox1.Text))
            cmd15.Parameters.AddWithValue("@fecha", Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", ""))
            cmd15.Parameters.AddWithValue("@ruc", Trim(TextBox30.Text))
            cmd15.Parameters.AddWithValue("@ncomercial", Trim(TextBox4.Text))
            cmd15.Parameters.AddWithValue("@GALGA", Convert.ToInt32(TextBox6.Text))
            cmd15.Parameters.AddWithValue("@DIAMETRO", Convert.ToInt32(TextBox5.Text))
            cmd15.Parameters.AddWithValue("@SISTEMAS", Convert.ToInt32(TextBox7.Text))
            cmd15.Parameters.AddWithValue("@AGUJAS", Convert.ToInt32(TextBox8.Text))
            cmd15.Parameters.AddWithValue("@densidad1", Convert.ToDouble(TextBox9.Text))
            cmd15.Parameters.AddWithValue("@ancho1", Convert.ToDouble(TextBox10.Text))
            cmd15.Parameters.AddWithValue("@densidad2", Convert.ToInt32(TextBox11.Text))
            cmd15.Parameters.AddWithValue("@ancho2", Convert.ToDouble(TextBox12.Text))
            cmd15.Parameters.AddWithValue("@densidad3", Convert.ToInt32(TextBox13.Text))
            cmd15.Parameters.AddWithValue("@ancho3", Convert.ToDouble(TextBox14.Text))
            cmd15.Parameters.AddWithValue("@lm1", Convert.ToDouble(TextBox15.Text))
            cmd15.Parameters.AddWithValue("@lm2", Convert.ToDouble(TextBox16.Text))
            cmd15.Parameters.AddWithValue("@lm3", Convert.ToDouble(TextBox17.Text))
            cmd15.Parameters.AddWithValue("@ruta1", Trim(TextBox18.Text))
            cmd15.Parameters.AddWithValue("@ruta2", Trim(TextBox19.Text))
            cmd15.Parameters.AddWithValue("@ruta3", Trim(TextBox20.Text))
            cmd15.Parameters.AddWithValue("@ruta4", Trim(TextBox21.Text))
            cmd15.Parameters.AddWithValue("@ruta5", Trim(TextBox22.Text))
            cmd15.Parameters.AddWithValue("@ruta6", Trim(TextBox23.Text))
            cmd15.Parameters.AddWithValue("@ruta7", Trim(TextBox24.Text))
            cmd15.Parameters.AddWithValue("@ruta8", Trim(TextBox25.Text))
            cmd15.Parameters.AddWithValue("@ruta9", Trim(TextBox26.Text))
            cmd15.Parameters.AddWithValue("@ruta10", Trim(TextBox27.Text))
            cmd15.Parameters.AddWithValue("@observacion", Trim(TextBox29.Text))
            cmd15.ExecuteNonQuery()

            Dim o, o1 As Integer
            o = DataGridView1.Rows.Count
            o1 = DataGridView2.Rows.Count
            For i = 0 To o - 1
                Dim cmd16 As New SqlCommand("EXEC custom_vianny.dbo.CaeSoft_FichaTextilTelaDetalleInsertUpdated  '01',@items,@codigotela,@codigohilo1,@codigohilo2,codigohilo3,'','',@cantidad,@codigohilo", conx)
                cmd16.Parameters.AddWithValue("@codigotela", Trim(TextBox1.Text))
                cmd16.Parameters.AddWithValue("@cantidad", (DataGridView1.Rows(i).Cells(4).Value))
                cmd16.Parameters.AddWithValue("@codigohilo", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd16.Parameters.AddWithValue("@codigohilo1", Mid(Trim(DataGridView1.Rows(i).Cells(2).Value), 3, 3))
                cmd16.Parameters.AddWithValue("@codigohilo2", Mid(Trim(DataGridView1.Rows(i).Cells(2).Value), 6, 1))
                cmd16.Parameters.AddWithValue("@codigohilo3", Mid(Trim(DataGridView1.Rows(i).Cells(2).Value), 7, 1))
                cmd16.Parameters.AddWithValue("@items", Trim(DataGridView1.Rows(i).Cells(0).Value))
                cmd16.ExecuteNonQuery()
            Next
            For i1 = 0 To o1 - 1
                Dim cmd17 As New SqlCommand("EXEC custom_vianny.dbo.Caesoft_TelaOPInsertUpdated '01' ,@codigotela, @op", conx)
                cmd17.Parameters.AddWithValue("@codigotela", Trim(TextBox1.Text))
                cmd17.Parameters.AddWithValue("@op", Trim(DataGridView2.Rows(i1).Cells(1).Value))
                cmd17.ExecuteNonQuery()
            Next
            MsgBox("SE GUARDO LA INFOMACION SOLICITADA")
            TabControl1.Enabled = False
            GroupBox2.Enabled = False
            GroupBox3.Enabled = False
            GroupBox4.Enabled = False
            GroupBox5.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            TextBox6.Enabled = False
            TextBox7.Enabled = False
            TextBox8.Enabled = False
            TextBox29.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = True
            TextBox1.Enabled = True
            limpiar()
        End If

    End Sub
    Sub limpiar()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox11.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        TextBox26.Text = ""
        TextBox27.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
    End Sub

End Class