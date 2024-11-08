Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class Nsalida
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Public comando As SqlCommand
    Dim ty, ty2, ty3, ty30, ty32 As New DataTable
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

            conn = New SqlDataAdapter("select dest_21s,rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where  Empr_21s =" + Trim(Label13.Text) + " AND almac_21s =" + alm + "order by nomb_21s", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "nomb_21s"
            ComboBox1.ValueMember = "dest_21s"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ty2.Clear()
        Dim sql1023 As String = " SELECT A.nomb_15m FROM custom_vianny.dbo.Mag1500 A  Where a.ccia ='" + Trim(Label13.Text) + "' and a.talm_15m ='" + Trim(ComboBox3.SelectedValue.ToString) + "'"
        Dim cmd1023 As New SqlCommand(sql1023, conx)
        Rsr24 = cmd1023.ExecuteReader()
        If Rsr24.Read() = True Then
            TextBox2.Text = Rsr24(0)

        End If
        Rsr24.Close()
        'abrir()
        'llenar_combo_box2(ComboBox1, ComboBox3.SelectedValue.ToString)


        TextBox4.Text = DateTimePicker1.Value.Month
        TextBox4.Select()

        TextBox4.Text = DateTimePicker1.Value.Month
        Select Case TextBox4.Text.Length
            Case "1" : TextBox4.Text = "0" & DateTimePicker1.Value.Month
            Case "2" : TextBox4.Text = TextBox4.Text

        End Select

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
    Dim Rsr24, Rsr2, Rsr25 As SqlDataReader
    Sub llenar_combo_alamcenes(ByVal ser As String)

        Try

            conn = New SqlDataAdapter("select talm_15m + '  |  '+ nomb_15m as nombre, talm_15m as almacen from custom_vianny.dbo.Mag1500  where ccia ='" + Trim(Label13.Text) + "' and flag_15m ='1' and (seriencr_15m ='" + ser + "' or seriencr_15m ='1_2' )", conx)
            conn.Fill(ty30)
            ComboBox3.DataSource = ty30
            ComboBox3.DisplayMember = "nombre"
            ComboBox3.ValueMember = "almacen"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Nsalida_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Label13.Text = "02" Then
            TextBox8.Text = "20459785834"
            TextBox10.Text = "GRAUS INDUSTRIAS TEXTIL"
        Else
            TextBox8.Text = "20508740361"
            TextBox10.Text = "CONSORCIO TEXTIL VIANNY"
        End If
        abrir()
        llenar_combo_alamcenes("1")
        Dim sql1023 As String = " SELECT A.nomb_15m FROM custom_vianny.dbo.Mag1500 A  Where a.ccia ='" + Trim(Label13.Text) + "' and a.talm_15m ='" + Trim(ComboBox3.SelectedValue.ToString) + "'"
        Dim cmd1023 As New SqlCommand(sql1023, conx)
        Rsr24 = cmd1023.ExecuteReader()
        If Rsr24.Read() = True Then
            TextBox2.Text = Rsr24(0)

        End If
        Rsr24.Close()
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        Button2.Enabled = False
        DataGridView1.Enabled = False
        PictureBox1.Enabled = False
        Button5.Enabled = False
        Button3.Enabled = False
        Button6.Enabled = False
        Button1.Enabled = False
        RadioButton1.Checked = True
        ComboBox2.SelectedIndex = 0
        TextBox4.Select()



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
                Select Case VAL.ToString.Length
                    Case "1" : DataGridView1.Rows(i1).Cells(0).Value = "00" & "" & VAL
                    Case "2" : DataGridView1.Rows(i1).Cells(0).Value = "0" & "" & VAL
                    Case "3" : DataGridView1.Rows(i1).Cells(0).Value = VAL
                End Select
            Next
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        If Trim(TextBox8.Text).Length = 0 Then
            MsgBox("FALTA INGRESAR EL CLIENTE")
        Else
            Dim art As New Articulos_stock
            ToolTip1.SetToolTip(PictureBox1, "AGREGAR PRODUCTOS")
            art.Label3.Text = ComboBox3.SelectedValue.ToString
            art.Label4.Text = 2
            art.Label5.Text = Label10.Text
            art.Label6.Text = Label13.Text
            art.Label7.Text = 2
            art.Label8.Text = Trim(TextBox4.Text) & Trim(TextBox5.Text)
            art.ShowDialog()
        End If


    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub



    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim JU As Integer
            Select Case TextBox4.Text.Length

                Case "1" : TextBox4.Text = "0" & "" & TextBox4.Text
                Case "2" : TextBox4.Text = TextBox4.Text
            End Select
            Dim ano As String
            ano = Convert.ToString(Year(DateTimePicker1.Value))
            abrir()
            TextBox5.Enabled = True
            enunciado = New SqlCommand("select top 1 ncom_3, substring(ncom_3,3,7) as ncom from custom_vianny.dbo.mag030f  where ccia_3 =" + Label13.Text + " and cperiodo_3 = '" + Label10.Text + "' and talm_3 = '" + ComboBox3.SelectedValue.ToString + "' " + "and ncom_3 like" + " '" + TextBox4.Text + "%' and ccom_3 = '2' order by ncom_3 desc", conx)
            respuesta = enunciado.ExecuteReader
            While respuesta.Read
                JU = respuesta.Item("ncom")
            End While
            respuesta.Close()
            TextBox5.Text = JU + 1
            Select Case TextBox5.Text.Length

                Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                Case "6" : TextBox5.Text = TextBox5.Text
            End Select

            TextBox5.ReadOnly = False
            TextBox5.Select()
            TextBox5.Focus()
        Else

            If e.KeyCode = Keys.F1 Then
                Ingresos_Salidas_almacen.Label3.Text = Label13.Text
                Ingresos_Salidas_almacen.Label4.Text = Label10.Text
                Ingresos_Salidas_almacen.Label5.Text = ComboBox3.SelectedValue.ToString
                Ingresos_Salidas_almacen.Label6.Text = Trim(TextBox4.Text)
                Ingresos_Salidas_almacen.Label7.Text = "4"
                Ingresos_Salidas_almacen.Show()
            End If
        End If
    End Sub
    Dim Rsr35, Rsr212, Rsr3001296 As SqlDataReader
    Sub correlativo()
        abrir()
        Dim JU1 As Integer
        Select Case TextBox4.Text.Length

            Case "1" : TextBox4.Text = "0" & "" & TextBox4.Text
            Case "2" : TextBox4.Text = TextBox4.Text
        End Select
        enunciado = New SqlCommand("select top 1 ncom_3, substring(ncom_3,3,7) as ncom from custom_vianny.dbo.mag030f  where ccia_3 =" + Label13.Text + " and cperiodo_3 = '" + Label10.Text + "' and talm_3 = '" + ComboBox3.SelectedValue.ToString + "' " + "and ncom_3 like" + " '" + TextBox4.Text + "%' and ccom_3 = '2' order by ncom_3 desc", conx)
        respuesta = enunciado.ExecuteReader
        While respuesta.Read
            JU1 = respuesta.Item("ncom")
        End While
        respuesta.Close()
        TextBox5.Text = JU1 + 1
        Select Case TextBox5.Text.Length

            Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
            Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
            Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
            Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
            Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
            Case "6" : TextBox5.Text = TextBox5.Text
        End Select

        TextBox5.ReadOnly = False
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
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, fila, suma As Integer
        suma = DataGridView1.Rows.Count
        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cant Consumo" Then
            abrir()
            Dim agregar As String = "delete  Almacen_Ficticio  WHERE  CodAfi = '" + Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) + "' AND AlmAfi ='" + Trim(ComboBox3.SelectedValue.ToString) + "' AND EmpAfi = '" + Trim(Label13.Text) + "' AND GuiAfi ='" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "' AND PerAfi='" + Trim(Label10.Text) + "'"
            ELIMINAR(agregar)

            For i10 = 0 To suma - 1

                If (Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString()) = Trim(DataGridView1.Rows(i10).Cells(2).Value.ToString())) Then


                    If (DataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString().Trim() = DataGridView1.Rows(i10).Cells(6).Value.ToString().Trim()) Then

                        Dim stock As Double
                        stock = 0
                        Dim sql1021 As String = "EXEC ReporteStockFisicoUnitario '" + Trim(Label13.Text) + "','" + Trim(Label10.Text) + "','" + Trim(ComboBox3.SelectedValue.ToString) + "','" + Trim(Label10.Text) + "0101','" + Trim(Label10.Text) + "1231','" + DataGridView1.Rows(i10).Cells(2).Value + "','" + DataGridView1.Rows(i10).Cells(2).Value + "',NULL, 2,'" + DataGridView1.Rows(i10).Cells(6).Value + "'"
                        Dim cmd1021 As New SqlCommand(sql1021, conx)
                        Rsr212 = cmd1021.ExecuteReader()

                        If Rsr212.Read() Then
                            stock = Convert.ToDouble(Rsr212(10))
                        End If
                        Rsr212.Close()
                        'MsgBox(stock.ToString)
                        Dim sql1022139 As String = "select SUM(CanAfi) from Almacen_Ficticio where CodAfi='" + Trim(DataGridView1.Rows(i10).Cells(2).Value) + "' and AlmAfi='" + Trim(ComboBox3.SelectedValue.ToString) + "' and ParAfi ='" + Trim(DataGridView1.Rows(i10).Cells(6).Value) + "'  group by CodAfi,AlmAfi,EmpAfi,ParAfi"
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
                        If Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(5).Value) > (stock - jo) Then
                            MsgBox("LA CANTIDAD SOLICITADA ES MAYOR AL STOCK")
                            DataGridView1.Rows(fila).Cells(5).Value = 0
                        Else
                            Dim cmd168 As New SqlCommand("insert into Almacen_Ficticio (ItmAfi,CodAfi,AlmAfi,EmpAfi,GuiAfi,PerAfi,CanAfi,ParAfi) values (@ItmAfi,@CodAfi,@AlmAfi,@EmpAfi,@GuiAfi,@PerAfi,@CanAfi,@ParAfi)", conx)
                            cmd168.Parameters.AddWithValue("@ItmAfi", Trim(DataGridView1.Rows(i10).Cells(0).Value))
                            cmd168.Parameters.AddWithValue("@CodAfi", Trim(DataGridView1.Rows(i10).Cells(2).Value))
                            cmd168.Parameters.AddWithValue("@AlmAfi", Trim(ComboBox3.SelectedValue.ToString))
                            cmd168.Parameters.AddWithValue("@EmpAfi", Trim(Label13.Text))
                            cmd168.Parameters.AddWithValue("@GuiAfi", Trim(TextBox4.Text) & Trim(TextBox5.Text))
                            cmd168.Parameters.AddWithValue("@PerAfi", Trim(Label10.Text))
                            cmd168.Parameters.AddWithValue("@CanAfi", Convert.ToDouble(DataGridView1.Rows(i10).Cells(5).Value))
                            cmd168.Parameters.AddWithValue("@ParAfi", Trim(DataGridView1.Rows(i10).Cells(6).Value))
                            cmd168.ExecuteNonQuery()
                        End If
                    End If
                End If
            Next
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Op" Then
            Dim sql10221348 As String = "SELECT count(ncom_3) as CANTIDAD FROM custom_vianny.dbo.qag0300 where ncom_3 ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) + "' and ccia ='" + Label13.Text + "'"
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

    Dim dt1, dt2, HG As New DataTable

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "INTERNA" And Trim(Label13.Text) = "01" Then
            TextBox8.Text = "20508740361"
            TextBox10.Text = "CONSORCIO TEXTIL VIANNY"
        Else
            If ComboBox2.Text = "INTERNA" And Trim(Label13.Text) = "02" Then
                TextBox8.Text = "20459785834"
                TextBox10.Text = "GRAUS INDUSTRIAS TEXTIL"
            End If
            TextBox8.ReadOnly = False
        End If
    End Sub
    Dim Rsr3008 As SqlDataReader
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub ToolTip2_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip2.Popup

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Public Sub limpiar()
        TextBox9.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox10.Text = ""
        TextBox1.Text = ""
        'TextBox17.Text = ""
        'TextBox19.Text = ""
    End Sub
    Dim Rsr300, Rsr215, Rsr216 As SqlDataReader
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim fun As New vinsertar_nia
        Dim func As New fnia
        Dim func1 As New insertardetallenia
        Dim func2 As New vnia
        Dim con As String
        Dim hn As New vnia
        Dim fg As New fnia


        Dim x, mj, valo, mp, mop, final As String
        Dim pp As Integer
        Dim VacioCliente As Boolean
        Dim ui, ct As Integer
        ct = 0
        ui = DataGridView1.Rows.Count
        For i11 = 0 To ui - 1
            If DataGridView1.Rows(i11).Cells(10).Value Is Nothing Then
                ct = ct + 1
            End If
        Next


        If String.IsNullOrEmpty(TextBox8.Text) Then
            VacioCliente = True
        End If
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
            If Trim(DataGridView1.Rows(i1).Cells(5).Value).Length = 0 Then
                Ccon = Ccon + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(6).Value).Length = 0 Then
                Cpartida = Cpartida + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(10).Value).Length = 0 Then
                Ccolor = Ccolor + 1
            End If
        Next

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

                                If ct > 0 Then
                                    MsgBox("FALTA AGREGAR LA CLASIFICACION DEL UN ARTICULO EN UNA FILA DE LA TABLA NSA")
                                Else

                                    x = Convert.ToString(TextBox1.Text.IndexOf("-", 0) + 1)

                                    If x > 0 Then

                                        valo = Trim(TextBox1.Text)
                                        Select Case x

                                            Case "1" : mj = "0000"
                                            Case "2" : mj = "000" & TextBox1.Text.Substring(0, x - 1)
                                            Case "3" : mj = "00" & TextBox1.Text.Substring(0, x - 1)
                                            Case "4" : mj = "0" & TextBox1.Text.Substring(0, x - 1)
                                            Case "5" : mj = TextBox1.Text.Substring(0, x - 1)

                                        End Select


                                        pp = Convert.ToInt32(TextBox1.Text.Length) - Convert.ToInt32(x)

                                        mp = TextBox1.Text.Substring(x, pp)
                                        Select Case mp.Length

                                            Case "1" : mop = "0000000" & mp
                                            Case "2" : mop = "000000" & mp
                                            Case "3" : mop = "00000" & mp
                                            Case "4" : mop = "0000" & mp
                                            Case "5" : mop = "000" & mp
                                            Case "6" : mop = "00" & mp
                                            Case "7" : mop = "0" & mp
                                            Case "8" : mop = mp
                                        End Select

                                        final = mj & "-" & mop
                                    Else

                                        final = Trim(TextBox1.Text)
                                    End If


                                    con = TextBox4.Text & TextBox5.Text
                                    hn.gcomprobante = con
                                    hn.galmacen = ComboBox3.SelectedValue.ToString
                                    hn.gncom = 2
                                    hn.gano = Label10.Text
                                    hn.gccia = Label13.Text
                                    fg.eliminar_nia(hn)
                                    Dim ns1 As New vnsacabece
                                    Dim ns2 As New fnsa

                                    ns1.gccia = Label13.Text
                                    ns1.gncom = con
                                    ns1.gfecha = DateTimePicker1.Value

                                    ns1.gorige_sali = ComboBox1.SelectedValue.ToString

                                    ns1.gglosa = Trim(TextBox9.Text)
                                    Select Case Trim(ComboBox4.Text)
                                        Case "GUIA" : ns1.gdoc = "009"
                                        Case "FACT" : ns1.gdoc = "001"
                                        Case "OTRO" : ns1.gdoc = "002"
                                        Case "" : ns1.gdoc = ""
                                    End Select

                                    ns1.gguia = final
                                    ns1.gsfactu = Mid(TextBox1.Text, 1, 4)
                                    ns1.gnfactu = Mid(TextBox1.Text, 6, 8)
                                    ns1.gruc = TextBox8.Text
                                    ns1.galmacen = ComboBox3.SelectedValue.ToString
                                    ns1.gusuario = TextBox7.Text
                                    ns1.gano = Label10.Text
                                    If ComboBox2.Text = "INTERNA" Then
                                        ns1.gtipointexte = "INT"
                                    Else
                                        ns1.gtipointexte = "EXT"
                                    End If
                                    Select Case ComboBox4.Text
                                        Case "GUIA" : ns1.gtdocento = "009"
                                        Case "FACT" : ns1.gtdocento = "001"
                                        Case "OTRO" : ns1.gtdocento = "002"
                                        Case "" : ns1.gtdocento = ""
                                    End Select
                                    ns1.gfase = TextBox11.Text
                                    ns1.gadevol = "0"
                                    ns1.gpedorig_3 = ""
                                    Dim i, num2 As Integer
                                    num2 = DataGridView1.Rows.Count


                                    If ns2.insertar_cabecera_nsa(ns1) Then
                                        For i = 0 To num2 - 1
                                            Dim nu As String
                                            nu = DataGridView1.Rows(i).Cells(0).Value
                                            Select Case nu.Length
                                                Case "1" : nu = "00" & "" & nu
                                                Case "2" : nu = "0" & "" & nu
                                                Case "3" : nu = nu
                                            End Select
                                            Dim ns3 As New vnsadetalle
                                            Dim ns4 As New fnsa
                                            Dim jj As New fingresopac
                                            Dim aa As New vpackingtela
                                            ns3.gccia = Label13.Text
                                            ns3.gncom = con
                                            ns3.gitem = nu
                                            ns3.glinea = DataGridView1.Rows(i).Cells(2).Value
                                            ns3.gcantidad = DataGridView1.Rows(i).Cells(5).Value
                                            ns3.gund = DataGridView1.Rows(i).Cells(8).Value
                                            If DataGridView1.Rows(i).Cells(1).Value = "" Then
                                                ns3.gop = ""
                                            Else
                                                ns3.gop = DataGridView1.Rows(i).Cells(1).Value
                                            End If
                                            ns3.gunidad_medidad = "RLL"
                                            ns3.gpartida = DataGridView1.Rows(i).Cells(6).Value
                                            ns3.grollo1 = DataGridView1.Rows(i).Cells(4).Value
                                            ns3.galmacen = ComboBox3.SelectedValue.ToString
                                            ns3.gordtejeduria = ""
                                            ns3.gano = Label10.Text
                                            If Convert.ToString(DataGridView1.Rows(i).Cells(6).Value) = "" Then
                                                ns3.glote = ""
                                            Else
                                                ns3.glote = DataGridView1.Rows(i).Cells(6).Value
                                            End If
                                            If Convert.ToString(DataGridView1.Rows(i).Cells(9).Value) = "0.00" Then
                                                ns3.gcantenvio = "0.00"
                                            Else
                                                ns3.gcantenvio = DataGridView1.Rows(i).Cells(9).Value
                                            End If
                                            Select Case Trim(DataGridView1.Rows(i).Cells(10).Value)
                                                Case "VERDE" : ns3.gclasif = "1"
                                                Case "AMARILLO" : ns3.gclasif = "2"
                                                Case "CELESTE" : ns3.gclasif = "3"
                                                Case "ROJO" : ns3.gclasif = "4"

                                            End Select
                                            If Convert.ToString(DataGridView1.Rows(i).Cells(11).Value) = "" Then
                                                ns3.gOcompra = ""
                                            Else
                                                ns3.gOcompra = DataGridView1.Rows(i).Cells(11).Value
                                            End If
                                            ns4.insertar_detalle_nsa(ns3)

                                            If Trim(DataGridView1.Rows(i).Cells(6).Value) = "" Then
                                                aa.gpartida = ""
                                            Else
                                                aa.gpartida = Trim(DataGridView1.Rows(i).Cells(6).Value)
                                            End If
                                            aa.galmacen = ComboBox3.SelectedValue.ToString
                                            aa.gseleccionado = 2
                                            'jj.actualizar_packing(aa)

                                            If Label15.Text = "1" Then
                                                Dim sum, suma As Double
                                                Dim cmd20 As New SqlCommand("update DetalleTelaManufactura  set EstDtm = '2'  where OpDtm =@op and CopDtm =@linea and CodDtm =@codigo", conx)
                                                cmd20.Parameters.AddWithValue("@op", (DataGridView1.Rows(i).Cells(1).Value.ToString.Trim))
                                                cmd20.Parameters.AddWithValue("@linea", (DataGridView1.Rows(i).Cells(2).Value.ToString.Trim))
                                                cmd20.Parameters.AddWithValue("@codigo", (DataGridView1.Rows(i).Cells(12).Value.ToString.Trim))
                                                cmd20.ExecuteNonQuery()

                                                Dim sql10 As String = "select ROUND(CASE WHEN cg.familia_17 = 'TPL' THEN cp.cxplineal * cantp_3 ELSE cp.kilos * q.cantp_3 END, 2) 
                                                                        from custom_vianny.dbo.Consumo_Op_DET cp
                                                                        LEFT JOIN custom_vianny.dbo.cag1700 cg ON cp.ccia = cg.ccia AND cp.tela = cg.linea_17 
                                                                        LEFT JOIN custom_vianny.dbo.qag0300 q ON cp.ccia = q.ccia AND cp.op = q.ncom_3 
                                                                        where op ='" + Trim(DataGridView1.Rows(i).Cells(1).Value.ToString.Trim) + "' and cp.ccia ='" + Label13.Text + "'"
                                                Dim cmd10 As New SqlCommand(sql10, conx)
                                                Rsr215 = cmd10.ExecuteReader()
                                                If Rsr215.Read() = True Then
                                                    sum = Rsr215(0)
                                                End If
                                                Rsr215.Close()

                                                Dim sql101 As String = "select sum(CanDtm) from DetalleTelaManufactura WHERE  OpDtm ='" + DataGridView1.Rows(i).Cells(1).Value.ToString.Trim + "'"
                                                Dim cmd101 As New SqlCommand(sql101, conx)
                                                Rsr216 = cmd101.ExecuteReader()
                                                If Rsr216.Read() = True Then
                                                    suma = Rsr216(0)
                                                End If
                                                Rsr216.Close()

                                                If sum = suma Then

                                                    Dim cmd21 As New SqlCommand("update RequerimientoTelaProd set EstRtp ='2' where IdRtp in (@IdRtp)", conx)
                                                    cmd21.Parameters.AddWithValue("@IdRtp", Trim(Label16.Text))
                                                    cmd21.ExecuteNonQuery()
                                                Else
                                                    Dim cmd211 As New SqlCommand("update RequerimientoTelaProd set EstRtp ='1' where IdRtp in (@IdRtp)", conx)
                                                    cmd211.Parameters.AddWithValue("@IdRtp", Trim(Label16.Text))
                                                    cmd211.ExecuteNonQuery()
                                                End If
                                            End If
                                        Next
                                        Dim agregar As String = "delete from Almacen_Ficticio where  AlmAfi='" + ComboBox3.SelectedValue.ToString + "' and EmpAfi='" + Label13.Text + "'  and GuiAfi ='" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "' "
                                        ELIMINAR(agregar)
                                        MessageBox.Show("Datos registrados correctamente ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                        Dim respuesta As DialogResult
                                        Dim hj2 As New flog
                                        Dim hj1 As New vlog
                                        hj1.gmodulo = "NSA-ALMACEN"
                                        hj1.gcuser = MDIParent1.Label3.Text
                                        If Label14.Text = "0" Then
                                            hj1.gaccion = 1
                                        Else
                                            hj1.gaccion = 2
                                        End If

                                        hj1.gpc = My.Computer.Name
                                        hj1.gfecha = DateTimePicker2.Value
                                        hj1.gdetalle = Trim(ComboBox3.SelectedValue.ToString) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
                                        hj1.gccia = Label13.Text
                                        hj2.insertar_log(hj1)
                                        respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                        If (respuesta = Windows.Forms.DialogResult.Yes) Then

                                            Reporte_Nia_Nsa.TextBox1.Text = ComboBox3.SelectedValue.ToString
                                            Reporte_Nia_Nsa.TextBox2.Text = 2
                                            Reporte_Nia_Nsa.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
                                            Reporte_Nia_Nsa.TextBox4.Text = Label10.Text
                                            Reporte_Nia_Nsa.TextBox5.Text = Label13.Text
                                            Reporte_Nia_Nsa.Show()
                                        End If
                                        limpiar()

                                        RadioButton1.Checked = True

                                    End If
                                    TextBox4.Enabled = False
                                    TextBox5.Enabled = True
                                    TextBox9.Enabled = False
                                    TextBox8.Enabled = False
                                    DateTimePicker1.Enabled = False
                                    DataGridView1.Enabled = False
                                    ComboBox1.Enabled = False
                                    ComboBox2.Enabled = False
                                    ComboBox4.Enabled = False
                                    DataGridView1.Rows.Clear()
                                    TextBox5.Select()
                                    Dim func12 As New fnsa
                                    Dim dts As New vnia
                                    TextBox4.Text = DateTime.Now.ToString("MM")

                                    dts.gmes = Me.TextBox4.Text
                                    dts.galmacen = ComboBox3.SelectedValue.ToString
                                    dts.gano = Label10.Text
                                    dts.gccia = Label13.Text
                                    Me.TextBox5.Text = func12.buscar_nsa(dts) + 1
                                    Select Case TextBox5.Text.Length

                                        Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                                        Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                                        Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                                        Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                                        Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                                        Case "6" : TextBox5.Text = TextBox5.Text
                                    End Select
                                End If
                            End If
                        End If
                    End If
                End If
            End If

        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'If Trim(TextBox4.Text) <> Month(DateTimePicker1.Value) Then
        '    MsgBox("LA FECHA DE EMISION NO CORRESPONDE AL MES DE REGISTRO")
        '    Dim me12 As String
        '    me12 = TextBox4.Text - Month(DateTimePicker1.Value)
        '    DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(me12)
        '    TextBox4.Select()
        'End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            ty2.Clear()
            abrir()
            llenar_combo_box2(ComboBox1, ComboBox3.SelectedValue.ToString)

            If Trim(TextBox4.Text) = "14" Then
                Dim me12 As String


                me12 = Mid(TextBox5.Text, 1, 2) - Month(DateTimePicker1.Value)

                DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(me12)
            Else
                Dim me12 As String
                me12 = TextBox4.Text - Month(DateTimePicker1.Value)
                DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(me12)
            End If


            Dim ml As New vnia
            Dim mk As New fnsa
            Dim i As Integer
            i = Len(TextBox5.Text)

            Select Case i
                Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                Case "6" : TextBox5.Text = TextBox5.Text

            End Select
            ml.gcomprobante = TextBox4.Text & TextBox5.Text

            ml.galmacen1 = ComboBox3.SelectedValue.ToString
            ml.gano = Label10.Text
            ml.gccia = Label13.Text
            If mk.mostrar_correlativo_nsa1(ml) = True Then
                Dim jk As New fnsa

                Button1.Enabled = True
                Button6.Enabled = True
                DataGridView1.Rows.Clear()
                dt1 = jk.mostrar_cabecera_nsa(ml)
                dt2 = jk.mostrar_detalle_nsa(ml)
                If dt1.Rows.Count <> 0 Then
                    DataGridView3.DataSource = dt1

                    TextBox9.Text = DataGridView3.Rows(0).Cells(0).Value
                    TextBox8.Text = DataGridView3.Rows(0).Cells(4).Value
                    If Convert.ToString(DataGridView3.Rows(0).Cells(6).Value) Is DBNull.Value Then
                        TextBox10.Text = ""
                    Else
                        TextBox10.Text = Convert.ToString(DataGridView3.Rows(0).Cells(6).Value)
                    End If

                    DateTimePicker1.Value = DataGridView3.Rows(0).Cells(1).Value
                    TextBox1.Text = DataGridView3.Rows(0).Cells(2).Value
                    If Trim(DataGridView3.Rows(0).Cells(9).Value) = "0" Then
                        CheckBox1.Checked = False
                    Else
                        CheckBox1.Checked = True
                    End If
                    abrir()
                    enunciado2 = New SqlCommand("select rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where  Empr_21s ='" + Trim(Label13.Text) + "' AND almac_21s='" + Trim(ComboBox3.SelectedValue.ToString) + "'  and dest_21s='" + Trim(DataGridView3.Rows(0).Cells(5).Value) + "'", conx)
                    respuesta2 = enunciado2.ExecuteReader
                    While respuesta2.Read
                        ComboBox1.Text = respuesta2.Item("nomb_21s")
                    End While
                    respuesta2.Close()
                    If Trim(DataGridView3.Rows(0).Cells(8).Value) = 1 Then
                        TextBox13.Text = ""
                        TextBox11.Text = ""
                    Else
                        TextBox11.Text = DataGridView3.Rows(0).Cells(8).Value
                        Select Case TextBox11.Text.Length
                            Case "1" : TextBox11.Text = "000" & "" & TextBox11.Text
                            Case "2" : TextBox11.Text = "00" & "" & TextBox11.Text
                            Case "3" : TextBox11.Text = "0" & "" & TextBox11.Text
                            Case "4" : TextBox11.Text = TextBox11.Text


                        End Select
                        enunciado3 = New SqlCommand(" SELECT  nomb_4 as Nombre FROM custom_vianny.dbo.Qag0400 A  Where CCIA = '" + Trim(Label13.Text) + "' and flag_4 ='1' and fase_4 ='" + Trim(TextBox11.Text) + "'", conx)
                        respuesta3 = enunciado3.ExecuteReader
                        While respuesta3.Read
                            TextBox13.Text = respuesta3.Item("Nombre")
                        End While
                        respuesta3.Close()
                    End If


                    'ComboBox4.Text = "GUIA"
                    If DataGridView3.Rows(0).Cells(7).Value = "EXT" Then
                        ComboBox2.SelectedIndex = 0
                    Else
                        ComboBox2.SelectedIndex = 1
                    End If
                    If DataGridView3.Rows(0).Cells(3).Value = 1 Then
                        RadioButton2.Checked = False
                        RadioButton1.Checked = True
                        RadioButton1.Enabled = True
                        Label12.Visible = False
                        Button4.Enabled = False
                    Else
                        RadioButton2.Checked = True
                        RadioButton1.Checked = False
                        RadioButton1.Enabled = False
                        Label12.Visible = True
                    End If

                End If
                If dt2.Rows.Count <> 0 Then
                    Dim nu1 As Integer
                    DataGridView4.DataSource = dt2
                    nu1 = DataGridView4.Rows.Count
                    DataGridView1.Rows.Add(nu1)
                    For i1 = 0 To nu1 - 1

                        DataGridView1.Rows(i1).Cells(0).Value = DataGridView4.Rows(i1).Cells(0).Value
                        DataGridView1.Rows(i1).Cells(1).Value = DataGridView4.Rows(i1).Cells(1).Value
                        DataGridView1.Rows(i1).Cells(2).Value = DataGridView4.Rows(i1).Cells(2).Value
                        DataGridView1.Rows(i1).Cells(3).Value = DataGridView4.Rows(i1).Cells(3).Value
                        DataGridView1.Rows(i1).Cells(4).Value = DataGridView4.Rows(i1).Cells(4).Value
                        DataGridView1.Rows(i1).Cells(5).Value = DataGridView4.Rows(i1).Cells(5).Value
                        DataGridView1.Rows(i1).Cells(6).Value = DataGridView4.Rows(i1).Cells(6).Value
                        DataGridView1.Rows(i1).Cells(8).Value = DataGridView4.Rows(i1).Cells(7).Value
                        DataGridView1.Rows(i1).Cells(9).Value = DataGridView4.Rows(i1).Cells(11).Value
                        DataGridView1.Rows(i1).Cells(11).Value = DataGridView4.Rows(i1).Cells(13).Value
                        Select Case Trim(DataGridView4.Rows(i1).Cells(12).Value)
                            Case "1" : DataGridView1.Rows(i1).Cells(10).Value = "VERDE"
                            Case "2" : DataGridView1.Rows(i1).Cells(10).Value = "AMARILLO"
                            Case "3" : DataGridView1.Rows(i1).Cells(10).Value = "CELESTE"
                            Case "4" : DataGridView1.Rows(i1).Cells(10).Value = "ROJO"
                            Case "" : DataGridView1.Rows(i1).Cells(10).Value = ""
                        End Select
                        Dim fh As New fguiasistema
                        Dim gb As New vguiacabecera
                        gb.gccia = Trim(Label13.Text)
                        gb.gCodArtIni = DataGridView4.Rows(i1).Cells(2).Value
                        gb.galmacen = ComboBox3.SelectedValue.ToString
                        If DataGridView1.Rows(i1).Cells(6).Value = "" Then
                            gb.gFiltroDescrip = "null"
                            gb.gModal = 1
                        Else
                            gb.gFiltroDescrip = DataGridView4.Rows(i1).Cells(6).Value
                            gb.gModal = 2
                        End If
                        gb.gano = Label10.Text
                        HG = fh.stock_guia(gb)
                        DataGridView5.DataSource = HG
                        If HG.Rows.Count > 0 Then
                            DataGridView1.Rows(i1).Cells(7).Value = DataGridView5.Rows(0).Cells(10).Value + DataGridView1.Rows(i1).Cells(5).Value
                        Else
                            DataGridView1.Rows(i1).Cells(7).Value = DataGridView1.Rows(i1).Cells(5).Value
                        End If


                    Next
                    DataGridView1.Columns(0).Frozen = True
                    DataGridView1.Columns(1).Frozen = True
                    DataGridView1.Columns(2).Frozen = True
                End If


            Else

                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                ComboBox3.Enabled = False
                ComboBox4.Enabled = True
                TextBox8.Enabled = True
                TextBox9.Enabled = True
                TextBox1.Enabled = True
                TextBox1.ReadOnly = False
                Button2.Enabled = True
                DataGridView1.Enabled = True
                PictureBox1.Enabled = True
                Button5.Enabled = True
                Button4.Enabled = True
                DataGridView1.Rows.Clear()
                TextBox9.Text = ""
                TextBox1.Text = ""
                TextBox8.Text = ""
                TextBox10.Text = ""
                TextBox8.Select()
                If Label15.Text = "1" Then
                    mostrar_automaticoguia()
                    ComboBox2.SelectedIndex = 0
                End If
            End If
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            ComboBox3.Enabled = False
            ComboBox4.SelectedIndex = -1
            DateTimePicker1.Select()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton2.Checked = True Then
            MsgBox("LA NOTA DE SALIADA ESTA ANULADA NO SE PUEDE MODIFICAR")
        Else
            Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label13.Text + "' AND A.talm_3 ='" + ComboBox3.SelectedValue.ToString + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
            Dim cmd102213 As New SqlCommand(sql102213, conx)
            Rsr300 = cmd102213.ExecuteReader()
            Dim jo As Integer
            If Rsr300.Read() Then
                jo = Rsr300(0)
            Else
                jo = 0
            End If

            Rsr300.Close()
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("QUIERES EDITAR ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                If jo = 0 Then
                    DataGridView1.Enabled = True
                    TextBox9.Enabled = True
                    DateTimePicker1.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox4.Enabled = True
                    TextBox4.Enabled = False
                    TextBox1.ReadOnly = False
                    TextBox8.ReadOnly = False
                    TextBox1.Enabled = True
                    TextBox5.Enabled = False
                    TextBox8.Enabled = True
                    Button2.Enabled = True
                    PictureBox1.Enabled = True
                    Button5.Enabled = True
                    Label14.Text = "1"
                    Button4.Enabled = True
                Else
                    MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
                End If

            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'System.Diagnostics.Process.Start("www.google.com")
        Dim sql1022138 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label13.Text + "' AND A.talm_3 ='" + ComboBox3.SelectedValue.ToString + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
        Dim cmd1022138 As New SqlCommand(sql1022138, conx)
        Rsr3008 = cmd1022138.ExecuteReader()
        Dim jo As Integer
        If Rsr3008.Read() Then
            jo = Rsr3008(0)
        Else
            jo = 0
        End If
        Rsr3008.Close()
        Dim respuesta As DialogResult
        Dim ml As New vnsacabece
        Dim mk As New fnsa
        ml.gncom = TextBox4.Text & TextBox5.Text
        ml.galmacen = ComboBox3.SelectedValue.ToString
        ml.gano = Label10.Text
        ml.gccia = Label13.Text
        respuesta = MessageBox.Show("QUIERES ANULAR NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then

            If jo = 0 Then

                Label12.Visible = True
                RadioButton2.Checked = True
                RadioButton1.Checked = False
                mk.anular_nSA(ml)

                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "NSA-ALMACEN"
                hj1.gcuser = MDIParent1.Label3.Text

                hj1.gaccion = 3


                hj1.gpc = My.Computer.Name
                hj1.gfecha = DateTimePicker2.Value
                hj1.gdetalle = Trim(ComboBox3.SelectedValue.ToString) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
                hj1.gccia = Label13.Text
                hj2.insertar_log(hj1)
                TextBox4.Enabled = True
                TextBox4.Select()
                DateTimePicker1.Enabled = False
                TextBox9.Enabled = False

                TextBox5.Enabled = True

                DataGridView1.Enabled = False
                Button5.Enabled = False
                Button6.Enabled = False
                Button4.Enabled = False
                Button1.Enabled = False

                Label12.Visible = False
                TextBox4.Enabled = False
                TextBox5.Enabled = True
                TextBox9.Enabled = False
                TextBox8.Enabled = False
                DateTimePicker1.Enabled = False
                DataGridView1.Enabled = False
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
                ComboBox4.Enabled = False
                DataGridView1.Rows.Clear()

                limpiar()
                TextBox5.Select()
                Dim func12 As New fnsa
                Dim dts As New vnia
                TextBox4.Text = DateTime.Now.ToString("MM")

                dts.gmes = Me.TextBox4.Text
                dts.galmacen = ComboBox3.SelectedValue.ToString
                dts.gano = Label10.Text
                dts.gccia = Label13.Text
                Me.TextBox5.Text = func12.buscar_nsa(dts) + 1
                Select Case TextBox5.Text.Length

                    Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                    Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                    Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                    Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                    Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                    Case "6" : TextBox5.Text = TextBox5.Text
                End Select
            Else
                MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
            End If

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        Button2.Enabled = False
        DataGridView1.Enabled = False
        DataGridView1.Rows.Clear()
        PictureBox1.Enabled = False
        Button5.Enabled = False
        Button3.Enabled = False
        Button6.Enabled = False
        Button1.Enabled = False
        RadioButton1.Checked = True
        limpiar()
        TextBox8.Enabled = False
        TextBox5.ReadOnly = False
        TextBox4.ReadOnly = False
        TextBox1.ReadOnly = False
        TextBox5.Enabled = True
        TextBox4.Enabled = True
        ComboBox3.Enabled = True
        TextBox4.Select()
        ty2.Clear()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ToolTip1.SetToolTip(Button6, "IMPRIMIR COMPROANTE")
        Reporte_Nia_Nsa.TextBox1.Text = ComboBox3.SelectedValue.ToString
        Reporte_Nia_Nsa.TextBox2.Text = 2
        Reporte_Nia_Nsa.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
        Reporte_Nia_Nsa.TextBox4.Text = Label10.Text
        Reporte_Nia_Nsa.TextBox5.Text = Label13.Text
        Reporte_Nia_Nsa.Show()
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown

        If e.KeyCode = Keys.F1 Then
            If Trim(TextBox5.Text).Length > 0 Then
                Dim cli As New Clientes
                cli.TextBox3.Text = "30"
                cli.ShowDialog()
            Else
                MessageBox.Show("FALTA INGRESAR EL CORRELATIVO DEL COMPROBANTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        End If
    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.F1 Then
            Procesos.Label1.Text = "3"
            Procesos.Label2.Text = Trim(Label13.Text)
            Procesos.Show()
        Else
            If e.KeyCode = Keys.Enter Then
                TextBox8.Select()
            End If
        End If
    End Sub

    Private Sub Nsalida_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Dim agregar As String = "delete from Almacen_Ficticio where  AlmAfi='" + ComboBox3.SelectedValue.ToString + "' and EmpAfi='" + Label13.Text + "'  and GuiAfi ='" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "' "
        ELIMINAR(agregar)
    End Sub

    Private Sub Nsalida_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        limpiar()
        DataGridView1.Rows.Clear()
        dt1.Rows.Clear()
        dt2.Rows.Clear()
    End Sub

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        If e.ColumnIndex = 4 Or e.ColumnIndex = 5 Then
            Dim cellValue As String = e.FormattedValue.ToString()
            If Not IsValidNumber(cellValue) Then
                MessageBox.Show("El valor ingresado no es válido. Debe ser un número entero o decimal con al menos un número antes del punto decimal.", "Error de Validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
            End If
        End If
    End Sub
    Private Function IsValidNumber(value As String) As Boolean
        Dim pattern As String = "^\d+(\.\d+)?$"
        Return System.Text.RegularExpressions.Regex.IsMatch(value, pattern)
    End Function

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox11.Select()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        NumConFrac(TextBox4, e)
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        NumConFrac(TextBox5, e)
    End Sub

    Sub validar_ingresos(columna As Integer, fila As Integer)
        Dim columnIndex As Integer = 5 ' Puedes cambiar este valor al índice de la columna que necesitas
        Dim columnIndex1 As Integer = 4
        Dim columnIndex2 As Integer = 6
        Dim columnIndex3 As Integer = 10
        If columna = columnIndex AndAlso fila >= 0 Then
            ' Activa el modo de edición de la celda
            DataGridView1.BeginEdit(True)
        End If
        If columna = columnIndex1 AndAlso fila >= 0 Then
            ' Activa el modo de edición de la celda
            DataGridView1.BeginEdit(True)
        End If
        If columna = columnIndex2 AndAlso fila >= 0 Then
            ' Activa el modo de edición de la celda
            DataGridView1.BeginEdit(True)
        End If
        If columna = columnIndex3 AndAlso fila >= 0 Then
            ' Activa el modo de edición de la celda
            DataGridView1.BeginEdit(True)
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        validar_ingresos(e.ColumnIndex, e.RowIndex)
    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        validar_ingresos(e.ColumnIndex, e.RowIndex)
    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim columnIndex2 As Integer = 1 ' Cambia este valor al índice de la columna que quieres validar
        Dim columnIndex As Integer = 5 ' Cambia este valor al índice de la columna que quieres validar
        Dim columnIndex1 As Integer = 4 ' Cambia este valor al índice de la columna que quieres validar
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex AndAlso TypeOf e.Control Is TextBox Then
            ' Agrega el controlador de eventos KeyPress al cuadro de texto de edición
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex1 AndAlso TypeOf e.Control Is TextBox Then
            ' Agrega el controlador de eventos KeyPress al cuadro de texto de edición
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex2 AndAlso TypeOf e.Control Is TextBox Then
            ' Agrega el controlador de eventos KeyPress al cuadro de texto de edición
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If
    End Sub

    Private Sub DataGridView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DataGridView1.KeyPress
        Dim currentCell As DataGridViewCell = DataGridView1.CurrentCell
        If currentCell.ColumnIndex = 5 Then
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
    Sub compras()
        DataGridView1.Enabled = True
        Button5.Enabled = True
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button6.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox4.Enabled = True
        Button1.Enabled = False
        Button2.Enabled = False

    End Sub

    Sub alamcenes(ByVal almacen As String)
        abrir()

        Try

            conn = New SqlDataAdapter("select talm_15m + '  |  '+ nomb_15m as nombre, talm_15m as almacen from custom_vianny.dbo.Mag1500  where ccia ='" + Trim(Label13.Text) + "' and  flag_15m ='1'  and talm_15m ='" + almacen + "'", conx)
            conn.Fill(ty32)
            ComboBox3.DataSource = ty32
            ComboBox3.DisplayMember = "nombre"
            ComboBox3.ValueMember = "almacen"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub EstablecerAlmacen(almacen As String)
        ' Verificar si el valor existe en el ComboBox antes de asignarlo
        alamcenes(almacen)
    End Sub
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
            DataGridView1.Rows(i5).Cells(5).Value = Rsr2(3)
            DataGridView1.Rows(i5).Cells(8).Value = Rsr2(4)
            DataGridView1.Rows(i5).Cells(6).Value = Rsr2(6).ToString().Trim()
            DataGridView1.Rows(i5).Cells(12).Value = Rsr2(5)
            DataGridView1.Rows(i5).Cells(10).Value = Rsr2(7).ToString().Trim()
            DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
            i5 = i5 + 1
        End While
        Rsr2.Close()

    End Sub
    'Sub Mostar_stock()
    '    'stock fisico

    '    For i8 = 0 To DataGridView1.Rows.Count - 1
    '        Dim sql1021 As String = "EXEC ReporteStockFisicoUnitario '" + Trim(Label30.Text) + "','" + Trim(Label29.Text) + "','" + Trim(Label23.Text) + "','" + Trim(Label29.Text) + "0101','" + Trim(Label29.Text) + "1231','" + DataGridView1.Rows(i8).Cells(2).Value + "','" + DataGridView1.Rows(i8).Cells(2).Value + "',NULL, 2,'" + DataGridView1.Rows(i8).Cells(6).Value + "'"
    '        Dim cmd1021 As New SqlCommand(sql1021, conx)
    '        Rsr25 = cmd1021.ExecuteReader()
    '        If Rsr25.Read() Then
    '            DataGridView1.Rows(i8).Cells(11).Value = Rsr25(10)
    '        End If
    '        Rsr25.Close()
    '    Next

    '    'fin stock
    'End Sub
End Class