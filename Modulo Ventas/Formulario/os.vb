Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class os
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim ty2, ty1 As New DataTable
    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("EXEC  CaeSoft_GetAllTipoOrdenesdeServicio", conx)
            conn.Fill(ty2)
            ComboBox3.DataSource = ty2
            ComboBox3.DisplayMember = "DESCRIPCION"
            ComboBox3.ValueMember = "CODIGO"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box1()
        Try

            conn = New SqlDataAdapter("SELECT cond_15k AS CON , nomb_15k AS NOMBRE FROM  custom_vianny.dbo.kag1500 where  ccia_15k ='" + Trim(Label16.Text) + "'", conx)
            conn.Fill(ty1)
            ComboBox1.DataSource = ty1
            ComboBox1.DisplayMember = "NOMBRE"
            ComboBox1.ValueMember = "CON"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub os_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim func As New ftcambio
        Dim dts As New vtcambio

        dts.gfecha = DateTimePicker1.Text
        ty2.Clear()
        ty1.Clear()
        abrir()
        llenar_combo_box2()

        RadioButton1.Checked = True
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = -1
        TextBox9.Text = func.mostrar_tipo_cambio(dts)
        TextBox2.Text = ""
        CheckBox2.Checked = True
        'CheckBox3.Checked = True
        TextBox3.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        ComboBox1.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        'DateTimePicker1.Enabled = False

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Trim(ComboBox3.Text) = "" Then

        Else
            TextBox1.Text = ComboBox3.SelectedValue.ToString
            abrir()
            Dim Rsr1991 As SqlDataReader
            Dim sql1011 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAlluLTIMONumeroOrdendeServicio '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "'"
            Dim cmd1011 As New SqlCommand(sql1011, conx)
            Rsr1991 = cmd1011.ExecuteReader()
            Rsr1991.Read()
            TextBox2.Text = Rsr1991(0)
            Rsr1991.Close()
            TextBox2.Select()
            Select Case ComboBox3.SelectedValue.ToString
                Case "C3" : TextBox14.Text = "0301"
                Case "D4" : TextBox14.Text = "0421"
                Case "E5" : TextBox14.Text = "0522"
                Case "H8" : TextBox14.Text = "0601"
                Case "I9" : TextBox14.Text = "0701"
            End Select
            Select Case ComboBox3.SelectedValue.ToString
                Case "C3" : TextBox15.Text = "CORTE"
                Case "D4" : TextBox15.Text = "CONFECCION"
                Case "E5" : TextBox15.Text = "ACABADO"
                Case "H8" : TextBox15.Text = "LAVANDERIA"
                Case "I9" : TextBox15.Text = "APLICACIONES Y PIEZAS"
            End Select

        End If




    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click


        'DataGridView1.Rows.Add()
        'Dim a = DataGridView1.Rows.Count
        ''DataGridView1.Rows(a - 1).Cells(0).Value = a
        'Select Case a.ToString.Length
        '    Case "1" : DataGridView1.Rows(a - 1).Cells(0).Value = "00" & "" & a + 1
        '    Case "2" : DataGridView1.Rows(a - 1).Cells(0).Value = "0" & "" & a + 1
        '    Case "3" : DataGridView1.Rows(a - 1).Cells(0).Value = a + 1

        'End Select
        DataGridView1.Rows.Add(1)

        Dim f, r As Integer
        f = DataGridView1.Rows.Count
        If f = 1 Then
            r = 1
            For i = 0 To f - 1
                DataGridView1.Rows(i).Cells(0).Value = 1

            Next
        Else
            r = 0
            For i = 1 To f - 1
                DataGridView1.Rows(i).Cells(0).Value = DataGridView1.Rows(i - 1).Cells(0).Value + 1 + r

            Next
        End If
        DataGridView1.Enabled = True

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
        Dim I18, VAL1 As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
            I18 = DataGridView1.Rows.Count
            For i1 = 0 To I18 - 1
                VAL1 = i1 + 1
                Select Case VAL1.ToString.Length
                    Case "1" : DataGridView1.Rows(i1).Cells(0).Value = "00" & "" & VAL1
                    Case "2" : DataGridView1.Rows(i1).Cells(0).Value = "0" & "" & VAL1
                    Case "3" : DataGridView1.Rows(i1).Cells(0).Value = VAL1
                End Select

                cant10 = VAL(DataGridView1.Rows(i1).Cells(12).Value)
                sum10 = cant10 + VAL(sum10)
                cant9 = VAL(DataGridView1.Rows(i1).Cells(13).Value)
                sum9 = cant9 + VAL(sum9)
                cant8 = VAL(DataGridView1.Rows(i1).Cells(14).Value)
                sum8 = cant8 + VAL(sum8)
            Next


            TextBox10.Text = sum10.ToString("N2")
            TextBox11.Text = sum9.ToString("N2")

            TextBox12.Text = sum8.ToString("N2")
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            Label18.Visible = True
            TextBox13.Visible = True
            TextBox13.Select()
        Else
            Label18.Visible = False
            TextBox13.Visible = False
        End If
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = "35556"
            Clientes.Show()
        End If
    End Sub
    Dim Rsr2 As SqlDataReader

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim VAL As Integer
            Select Case Trim(TextBox13.Text).Length
                Case "1" : TextBox13.Text = "01" & "0000000" & TextBox13.Text
                Case "2" : TextBox13.Text = "01" & "000000" & TextBox13.Text
                Case "3" : TextBox13.Text = "01" & "00000" & TextBox13.Text
                Case "4" : TextBox13.Text = "01" & "0000" & TextBox13.Text
                Case "5" : TextBox13.Text = "01" & "000" & TextBox13.Text
                Case "6" : TextBox13.Text = "01" & "00" & TextBox13.Text
                Case "7" : TextBox13.Text = "01" & "0" & TextBox13.Text
                Case "8" : TextBox13.Text = "01" & TextBox13.Text
                Case "9" : TextBox13.Text = "0" & TextBox13.Text
            End Select
            Dim sql102 As String = "select ncom_3,linea_3,ISNULL(a.nomb_17,Q.descri_3),ISNULL( a.unid_17,'UND'),program_3 from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + Trim(TextBox13.Text) + "' and q.ccia ='" + Trim(Label16.Text) + "'
"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            Dim i5 As Integer
            i5 = DataGridView1.Rows.Count

            If Rsr2.Read() = True Then
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i5).Cells(1).Value = Rsr2(4)
                DataGridView1.Rows(i5).Cells(2).Value = Rsr2(0)
                DataGridView1.Rows(i5).Cells(4).Value = Rsr2(1)
                DataGridView1.Rows(i5).Cells(5).Value = Rsr2(2)
                DataGridView1.Rows(i5).Cells(6).Value = Rsr2(3)
                DataGridView1.Rows(i5).Cells(7).Value = 0
                DataGridView1.Rows(i5).Cells(8).Value = 0
                DataGridView1.Rows(i5).Cells(9).Value = 0
                DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                VAL = DataGridView1.Rows(i5).Cells(0).Value
                Select Case VAL.ToString.Length
                    Case "1" : DataGridView1.Rows(i5).Cells(0).Value = "00" & "" & VAL
                    Case "2" : DataGridView1.Rows(i5).Cells(0).Value = "0" & "" & VAL
                    Case "3" : DataGridView1.Rows(i5).Cells(0).Value = VAL
                End Select
                Rsr2.Close()
                TextBox13.Text = ""
            Else
                Rsr2.Close()
                MsgBox("LA OP INGRESADA NO EXISTE")
            End If
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
        'If Label20.Text = "0" Then
        If CheckBox2.Checked = False Then

                Dim i8 As Integer
                i8 = DataGridView1.RowCount
                CheckBox2.Enabled = True
            CheckBox3.Checked = False
            CheckBox3.Enabled = False
            For i = 0 To i8 - 1

                    DataGridView1.Rows(i).Cells(14).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(8).Value
                    DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(8).Value
                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(14).DefaultCellStyle.Format = "n2"
                    DataGridView1.Rows(i).Cells(13).Value = "0.00"

                Next

                For a9 = 0 To i8 - 1
                    cant10 = Val(DataGridView1.Rows(a9).Cells(12).Value)
                    sum10 = cant10 + Val(sum10)
                    cant9 = Val(DataGridView1.Rows(a9).Cells(13).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(DataGridView1.Rows(a9).Cells(14).Value)
                    sum8 = cant8 + Val(sum8)

                Next
                TextBox10.Text = sum10.ToString("N2")
                TextBox11.Text = sum9.ToString("N2")

                TextBox12.Text = sum8.ToString("N2")
            Else
            If CheckBox2.Checked = True Then

                CheckBox3.Enabled = True
                Dim i8 As Integer
                i8 = DataGridView1.RowCount
                For i = 0 To i8 - 1

                    DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(8).Value

                    DataGridView1.Rows(i).Cells(14).Value = DataGridView1.Rows(i).Cells(9).Value * 1.18
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView1.Rows(i).Cells(14).Value - DataGridView1.Rows(i).Cells(12).Value
                    Me.DataGridView1.Columns(14).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                Next

                For a9 = 0 To i8 - 1
                    cant10 = Val(DataGridView1.Rows(a9).Cells(12).Value)
                    sum10 = cant10 + Val(sum10)
                    cant9 = Val(DataGridView1.Rows(a9).Cells(13).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(DataGridView1.Rows(a9).Cells(14).Value)
                    sum8 = cant8 + Val(sum8)

                Next
                TextBox10.Text = sum10.ToString("N2")
                TextBox11.Text = sum9.ToString("N2")

                TextBox12.Text = sum8.ToString("N2")
            End If

        End If

        'End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        'If Label20.Text = "0" Then
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
            If CheckBox3.Checked = True Then
                Dim i8 As Integer
                i8 = DataGridView1.RowCount
                For i = 0 To i8 - 1
                    DataGridView1.Rows(i).Cells(14).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(8).Value
                    DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(14).Value / 1.18
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView1.Rows(i).Cells(14).Value - DataGridView1.Rows(i).Cells(12).Value
                    Me.DataGridView1.Columns(14).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"

                Next

                For a9 = 0 To i8 - 1
                    cant10 = Val(DataGridView1.Rows(a9).Cells(12).Value)
                    sum10 = cant10 + Val(sum10)
                    cant9 = Val(DataGridView1.Rows(a9).Cells(13).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(DataGridView1.Rows(a9).Cells(14).Value)
                    sum8 = cant8 + Val(sum8)

                Next
                TextBox10.Text = sum10.ToString("N2")
                TextBox11.Text = sum9.ToString("N2")

                TextBox12.Text = sum8.ToString("N2")
            Else
                If CheckBox3.Checked = False And CheckBox2.Checked = True Then

                    Dim i81 As Integer
                    i81 = DataGridView1.RowCount
                    For i = 0 To i81 - 1

                        DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(8).Value

                        DataGridView1.Rows(i).Cells(14).Value = DataGridView1.Rows(i).Cells(14).Value * 1.18
                        DataGridView1.Rows(i).Cells(13).Value = DataGridView1.Rows(i).Cells(14).Value - DataGridView1.Rows(i).Cells(12).Value
                        Me.DataGridView1.Columns(14).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                    Next

                    For a9 = 0 To i81 - 1
                        cant10 = Val(DataGridView1.Rows(a9).Cells(12).Value)
                        sum10 = cant10 + Val(sum10)
                        cant9 = Val(DataGridView1.Rows(a9).Cells(13).Value)
                        sum9 = cant9 + Val(sum9)
                        cant8 = Val(DataGridView1.Rows(a9).Cells(14).Value)
                        sum8 = cant8 + Val(sum8)

                    Next
                    TextBox10.Text = sum10.ToString("N2")
                    TextBox11.Text = sum9.ToString("N2")

                    TextBox12.Text = sum8.ToString("N2")
                End If
            End If
        'End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, a9, fila As Integer
        Dim cant10, sum10, cant9, sum9, cant8, sum8, cant7, sum7 As Double
        i = DataGridView1.Rows.Count
        fila = e.RowIndex

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cantidad" Then
            Try
                'DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                'DataGridView1.Rows(fila).Cells(14).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                'DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(9).Value / 1.18
                'DataGridView1.Rows(fila).Cells(13).Value = DataGridView1.Rows(fila).Cells(9).Value - DataGridView1.Rows(fila).Cells(12).Value
                'Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                'Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                'Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                If CheckBox3.Checked = False And CheckBox2.Checked = True Then

                    DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)

                    DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(9).Value
                    DataGridView1.Rows(fila).Cells(14).Value = Val(DataGridView1.Rows(fila).Cells(9).Value * 1.18)
                    DataGridView1.Rows(fila).Cells(13).Value = DataGridView1.Rows(fila).Cells(14).Value - DataGridView1.Rows(fila).Cells(12).Value
                    Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                Else
                    If CheckBox3.Checked = True And CheckBox2.Checked = True Then
                        DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                        DataGridView1.Rows(fila).Cells(14).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                        DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(9).Value / 1.18
                        DataGridView1.Rows(fila).Cells(13).Value = DataGridView1.Rows(fila).Cells(9).Value - DataGridView1.Rows(fila).Cells(12).Value
                        Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"

                    Else
                        If CheckBox3.Checked = False And CheckBox2.Checked = False Then
                            DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(14).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(12).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(13).Value = 0
                            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                        End If
                    End If
                End If
                DataGridView1.Focus()
                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(8)
                DataGridView1.BeginEdit(True)
            Catch ex As Exception
                MsgBox("NO SELECCIONO UNA FILA CORRECTA")
            End Try

        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "P.Unitario" Then
            Try
                'DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                'DataGridView1.Rows(fila).Cells(14).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                'DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(9).Value / 1.18
                'DataGridView1.Rows(fila).Cells(13).Value = DataGridView1.Rows(fila).Cells(9).Value - DataGridView1.Rows(fila).Cells(12).Value
                'Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                'Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                'Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                If CheckBox3.Checked = False And CheckBox2.Checked = True Then

                    DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)

                    DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(9).Value
                    DataGridView1.Rows(fila).Cells(14).Value = Val(DataGridView1.Rows(fila).Cells(9).Value * 1.18)
                    DataGridView1.Rows(fila).Cells(13).Value = DataGridView1.Rows(fila).Cells(14).Value - DataGridView1.Rows(fila).Cells(12).Value
                    Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                Else
                    If CheckBox3.Checked = True And CheckBox2.Checked = True Then
                        DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                        DataGridView1.Rows(fila).Cells(14).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                        DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(9).Value / 1.18
                        DataGridView1.Rows(fila).Cells(13).Value = DataGridView1.Rows(fila).Cells(9).Value - DataGridView1.Rows(fila).Cells(12).Value
                        Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"

                    Else
                        If CheckBox3.Checked = False And CheckBox2.Checked = False Then
                            DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(14).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(12).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(13).Value = 0
                            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                        End If
                    End If
                End If
                'DataGridView1.Focus()
                'DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(7)
                'DataGridView1.BeginEdit(True)


            Catch ex As Exception
                MsgBox("NO SELECCIONO UNA FILA CORRECTA")
            End Try

        End If
        For a9 = 0 To i - 1
            cant10 = Val(DataGridView1.Rows(a9).Cells(12).Value)
            sum10 = cant10 + Val(sum10)
            cant9 = Val(DataGridView1.Rows(a9).Cells(13).Value)
            sum9 = cant9 + Val(sum9)
            cant8 = Val(DataGridView1.Rows(a9).Cells(14).Value)
            sum8 = cant8 + Val(sum8)


        Next
        TextBox10.Text = sum10.ToString("N2")
        TextBox11.Text = sum9.ToString("N2")

        TextBox12.Text = sum8.ToString("N2")
    End Sub
    Dim Rsr20, Rsr21, Rsr22 As SqlDataReader
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            llenar_combo_box1()
            TextBox3.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            ComboBox1.Enabled = True
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True

            Button2.Enabled = True
            Button3.Enabled = True

            Button6.Enabled = True
            'DateTimePicker1.Enabled = True
            DataGridView1.Enabled = True
            TextBox3.Select()
            Button1.Enabled = True
            Dim i As Integer
            i = Len(TextBox2.Text)
            Select Case i
                Case "1" : TextBox2.Text = "00000" & "" & TextBox2.Text
                Case "2" : TextBox2.Text = "0000" & "" & TextBox2.Text
                Case "3" : TextBox2.Text = "000" & "" & TextBox2.Text
                Case "4" : TextBox2.Text = "00" & "" & TextBox2.Text
                Case "5" : TextBox2.Text = "0" & "" & TextBox2.Text
                Case "6" : TextBox2.Text = TextBox2.Text

            End Select
            abrir()

            Dim sql10215 As String = "select COUNT(ncom_4) from custom_vianny.dbo.lap0400 p inner join custom_vianny.dbo.laG0400 g on p.ccia_4a =g.ccia_4 and p.ncom_4a = g.ncom_4  where ncom_4 ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' and ccia_4 ='" + Trim(Label16.Text) + "'"
            Dim cmd10215 As New SqlCommand(sql10215, conx)
            Rsr22 = cmd10215.ExecuteReader()
            Rsr22.Read()

            If Rsr22(0) = "0" Then
                'Label20.Text = 0
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = ""
                TextBox8.Text = ""
                TextBox10.Text = ""
                TextBox11.Text = ""
                TextBox12.Text = ""
                DataGridView1.Rows.Clear()
                DataGridView1.Enabled = True
                Rsr22.Close()
            Else
                Button5.Enabled = True
                DataGridView1.Rows.Clear()
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                TextBox6.Enabled = False
                TextBox7.Enabled = False
                TextBox8.Enabled = False
                TextBox10.Enabled = False
                TextBox11.Enabled = False
                TextBox12.Enabled = False
                DataGridView1.Enabled = False
                Button4.Enabled = True
                ComboBox1.Enabled = False
                TextBox14.Enabled = False
                'Label20.Text = 1

                Rsr22.Close()

                Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAllOrdenesdeServicioCabecera '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "','" + Trim(TextBox2.Text) + "'"
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr21 = cmd1021.ExecuteReader()
                Dim i51 As Integer
                i51 = 0
                While Rsr21.Read()

                    TextBox3.Text = Rsr21(2)
                    TextBox4.Text = Rsr21(24)
                    TextBox5.Text = Rsr21(5)
                    TextBox6.Text = Rsr21(8)
                    TextBox7.Text = Trim(Rsr21(9))
                    TextBox8.Text = Trim(Rsr21(12))
                    TextBox14.Text = Rsr21(27)
                    TextBox15.Text = Rsr21(28)
                    TextBox9.Text = Rsr21(6)
                    If Trim(Rsr21(14)) = "2" Then
                        Label21.Visible = True
                    Else
                        Label21.Visible = False
                    End If
                    If Rsr21(35) = "1" Then
                        CheckBox1.Checked = True
                    Else
                        CheckBox1.Checked = False
                    End If
                    If Rsr21(17) = "1" Then

                        TextBox10.Text = Rsr21(18)
                        TextBox11.Text = Rsr21(19)
                        TextBox12.Text = Rsr21(20)
                    Else
                        If Rsr21(17) = "2" Then
                            TextBox10.Text = Rsr21(21)
                            TextBox11.Text = Rsr21(22)
                            TextBox12.Text = Rsr21(23)
                        End If

                    End If

                    DateTimePicker1.Value = Rsr21(3)
                    DateTimePicker2.Value = Rsr21(4)
                    If Rsr21(14) = 1 Then
                        RadioButton1.Checked = True
                    Else
                        RadioButton2.Checked = True
                    End If
                    If Rsr21(15) = "1" Then
                        CheckBox2.Checked = True
                    Else
                        CheckBox2.Checked = False
                    End If
                    If Rsr21(16) = "1" Then
                        CheckBox3.Checked = True
                    Else
                        CheckBox3.Checked = False
                    End If
                    If Rsr21(17) = "1" Then
                        ComboBox2.SelectedIndex = 0
                    Else
                        ComboBox2.SelectedIndex = 1
                    End If

                    abrir()
                    enunciado2 = New SqlCommand("select nomb_15k as nomb_15k from custom_vianny.dbo.kag1500 where  ccia_15k ='" + Trim(Label16.Text) + "' AND cond_15k='" + Trim(Rsr21(7)) + "'", conx)
                    respuesta2 = enunciado2.ExecuteReader
                    While respuesta2.Read

                        ComboBox1.Text = respuesta2.Item("nomb_15k")
                    End While
                    respuesta2.Close()




                End While
                Rsr21.Close()

                Dim sql10212 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAllOrdenesdeServicioDetalle '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "','" + Trim(TextBox2.Text) + "'"
                Dim cmd10212 As New SqlCommand(sql10212, conx)
                Rsr20 = cmd10212.ExecuteReader()


                While Rsr20.Read()
                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(i51).Cells(0).Value = Rsr20(3)
                    'Dim i4 As Integer
                    'i4 = Len(Rsr20(3))

                    'Select Case i4
                    '    Case "1" : DataGridView1.Rows(i51).Cells(0).Value = "00" & "" & Convert.ToString(Rsr20(3))
                    '    Case "2" : DataGridView1.Rows(i51).Cells(0).Value = "0" & "" & Convert.ToString(Rsr20(3))
                    '    Case "3" : DataGridView1.Rows(i51).Cells(0).Value = Convert.ToString(Rsr20(3))
                    'End Select
                    DataGridView1.Rows(i51).Cells(1).Value = Rsr20(18)
                    DataGridView1.Rows(i51).Cells(2).Value = Rsr20(2)
                    DataGridView1.Rows(i51).Cells(4).Value = Rsr20(4)
                    DataGridView1.Rows(i51).Cells(5).Value = Rsr20(6)
                    DataGridView1.Rows(i51).Cells(6).Value = Rsr20(7)
                    DataGridView1.Rows(i51).Cells(7).Value = Rsr20(8)
                    DataGridView1.Rows(i51).Cells(11).Value = Trim(Rsr20(13))
                    DataGridView1.Rows(i51).Cells(15).Value = Trim(Rsr20(14))
                    DataGridView1.Rows(i51).Cells(16).Value = Trim(Rsr20(15))
                    DataGridView1.Rows(i51).Cells(17).Value = Trim(Rsr20(16))
                    DataGridView1.Rows(i51).Cells(18).Value = Trim(Rsr20(17))
                    If ComboBox2.Text = "SOLES" Then
                        DataGridView1.Rows(i51).Cells(8).Value = Rsr20(9)
                        DataGridView1.Rows(i51).Cells(9).Value = Rsr20(10)
                    Else
                        DataGridView1.Rows(i51).Cells(8).Value = Rsr20(11)
                        DataGridView1.Rows(i51).Cells(9).Value = Rsr20(12)
                    End If
                    Dim cant10, sum10, cant9, sum9, cant8, sum8, cant101, sum101, cant91, sum91, cant81, sum81, cant102, sum102, cant92, sum92, cant82, sum82 As Double
                    If CheckBox2.Checked = True And CheckBox3.Checked = False Then
                        Dim i8 As Integer
                        i8 = DataGridView1.RowCount
                        CheckBox2.Enabled = True
                        CheckBox3.Checked = False
                        For i = 0 To i8 - 1
                            DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(8).Value

                            DataGridView1.Rows(i).Cells(14).Value = DataGridView1.Rows(i).Cells(9).Value * 1.18
                            DataGridView1.Rows(i).Cells(13).Value = DataGridView1.Rows(i).Cells(14).Value - DataGridView1.Rows(i).Cells(12).Value

                            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(14).DefaultCellStyle.Format = "n2"


                        Next
                        'For a9 = 0 To i8 - 1
                        cant10 = Val(DataGridView1.Rows(i51).Cells(12).Value)
                        sum10 = cant10 + Val(sum10)
                        cant9 = Val(DataGridView1.Rows(i51).Cells(13).Value)
                        sum9 = cant9 + Val(sum9)
                        cant8 = Val(DataGridView1.Rows(i51).Cells(14).Value)
                        sum8 = cant8 + Val(sum8)

                        'Next
                        TextBox10.Text = sum10.ToString("N2")
                        TextBox11.Text = sum9.ToString("N2")

                        TextBox12.Text = sum8.ToString("N2")
                    Else
                        If CheckBox2.Checked = False Then
                            Dim i9 As Integer
                            i9 = DataGridView1.RowCount
                            For i = 0 To i9 - 1
                                DataGridView1.Rows(i).Cells(14).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(8).Value
                                DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(8).Value
                                DataGridView1.Rows(i).Cells(13).Value = "0.00"
                                Me.DataGridView1.Columns(14).DefaultCellStyle.Format = "n2"
                                Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                                Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                            Next
                            'For a91 = 0 To i9 - 1
                            cant101 = Val(DataGridView1.Rows(i51).Cells(12).Value)
                            sum101 = cant101 + Val(sum101)
                            cant91 = Val(DataGridView1.Rows(i51).Cells(13).Value)
                            sum91 = cant91 + Val(sum91)
                            cant81 = Val(DataGridView1.Rows(i51).Cells(14).Value)
                            sum81 = cant81 + Val(sum81)

                            'Next
                            TextBox10.Text = sum101.ToString("N2")
                            TextBox11.Text = sum91.ToString("N2")

                            TextBox12.Text = sum81.ToString("N2")
                        Else
                            If CheckBox2.Checked = True And CheckBox3.Checked = True Then
                                Dim i10 As Integer
                                i10 = DataGridView1.RowCount

                                For i = 0 To i10 - 1
                                    DataGridView1.Rows(i).Cells(14).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(8).Value

                                    DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(9).Value / 1.18
                                    DataGridView1.Rows(i).Cells(13).Value = DataGridView1.Rows(i).Cells(9).Value - DataGridView1.Rows(i).Cells(12).Value

                                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                                    Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                                    Me.DataGridView1.Columns(14).DefaultCellStyle.Format = "n2"
                                Next
                                'For a92 = 0 To i10 - 1
                                'MsgBox(Val(DataGridView1.Rows(i51).Cells(12).Value))
                                cant102 = Val(DataGridView1.Rows(i51).Cells(12).Value)
                                sum102 = cant102 + Val(sum102)
                                cant92 = Val(DataGridView1.Rows(i51).Cells(13).Value)
                                sum92 = cant92 + Val(sum92)
                                cant82 = Val(DataGridView1.Rows(i51).Cells(14).Value)
                                sum82 = cant82 + Val(sum82)

                                'Next
                                TextBox10.Text = sum102.ToString("N2")
                                TextBox11.Text = sum92.ToString("N2")

                                TextBox12.Text = sum82.ToString("N2")
                            End If
                        End If


                    End If
                    i51 = i51 + 1
                End While

                Rsr20.Close()
            End If

            If Trim(ComboBox3.Text) = "C3    | CORTE" Then
                DateTimePicker2.Value = DateTimePicker1.Value.AddDays(2)
            Else
                If Trim(ComboBox3.Text) = "D4    | CONFECCION" Then
                    DateTimePicker2.Value = DateTimePicker1.Value.AddDays(8)
                Else
                    If Trim(ComboBox3.Text) = "E5    | ACABADO" Then
                        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(2)
                    Else
                        If Trim(ComboBox3.Text) = "H8    | LAVANDERIA" Then
                            DateTimePicker2.Value = DateTimePicker1.Value.AddDays(3)
                        Else
                            DateTimePicker2.Value = DateTimePicker1.Value
                        End If
                    End If
                End If
            End If

        Else
            If e.KeyCode = Keys.F1 Then
                Det_Os.Label4.Text = Trim(Label16.Text)
                Det_Os.Label5.Text = TextBox1.Text
                Det_Os.Label6.Text = "1"
                Det_Os.Show()
            End If
        End If




    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FORM_OS.TextBox1.Text = TextBox1.Text & TextBox2.Text
        FORM_OS.TextBox2.Text = TextBox1.Text & TextBox2.Text
        FORM_OS.TextBox3.Text = Label16.Text
        FORM_OS.TextBox4.Text = "\\172.21.0.1\erpcaesoft$\LIB.APPSV2\imagenes\"

        FORM_OS.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If CheckBox1.Checked = True Then
            MsgBox("NO SE PUEDE MODIFICAR LA ORDEN DE SERVICIO QUE ESTA APROBADA")
        Else
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            TextBox10.Enabled = True
            TextBox11.Enabled = True
            TextBox12.Enabled = True
            DataGridView1.Enabled = True
            ComboBox1.Enabled = True
            TextBox14.Enabled = True
            Label22.Text = "1"
        End If


        DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
        DataGridView1.BeginEdit(True)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cmd15 As New SqlCommand("IF NOT EXISTS( SELECT NCOM_4 FROM custom_vianny.dbo.LAG0400 WHERE CCIA_4 = '" + Label16.Text + "' AND NCOM_4 = '" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' )
				INSERT INTO custom_vianny.dbo.LAG0400 ( CCIA_4 , ncom_4 , fich_4 , fcom_4 , fcome_4 , cond_4 , rubr_4 ,tcam_4 , mone_4 , tot1_4 , tot2_4 , vvta1_4 , vvta2_4 , igv1_4 , igv2_4 , cigv_4 , refe_4 ,gloa_4 , glob_4 , gloc_4 , glod_4 , aigv_4 , talm_4 , tipoc_4  , ordenp_4 , lugar_4 ,  subfase_4,flag_4  , cerrado_4 , dpiespag_4 , atencion_4, detraccion_4, coddetrac_4, Aprobado_4, userapro_4, Tipo_4 ) VALUES
				                                     (@CCIA_4 ,@ncom_4 ,@fich_4 ,@fcom_4 ,@fcome_4 ,@cond_4 , ''     ,@tcam_4,@mone_4 ,@tot1_4 ,@tot2_4 ,@vvta1_4 ,@vvta2_4 ,@igv1_4 ,@igv2_4 ,@cigv_4 , ''     ,@gloa_4, ''     , ''     , ''     ,@aigv_4 , ''     , ''       , ''       ,@lugar_4 , @subfase_4,@flag_4 ,@cerrado_4 , ''         , ''        , 0           , ''         , 0         , ''        , '01')
		ELSE
		   UPDATE custom_vianny.dbo.LAG0400 SET fich_4 =@fich_4,
							    fcom_4  = @fcom_4  ,
							    fcome_4 = @fcome_4 ,
								cond_4  = @cond_4 ,
								rubr_4  = '' ,
								tcam_4  = @tcam_4 ,
								mone_4  = @mone_4 ,
								tot1_4  = @tot1_4 ,
								tot2_4  = @tot2_4 ,
								vvta1_4 = @vvta1_4 ,
								vvta2_4 = @vvta2_4 ,
								igv1_4  = @igv1_4 ,
								igv2_4  = @igv2_4 ,			
								cigv_4  = @cigv_4  ,
								refe_4  = '',
								gloa_4  = @gloa_4,
								glob_4  = '',
								gloc_4  = '',
								glod_4  = '',
								aigv_4  = @aigv_4,
								talm_4  = '',
								tipoc_4 = '',
								ordenp_4= '',
								lugar_4 = @lugar_4,
								subfase_4= @subfase_4,
								flag_4    = @flag_4,
								cerrado_4 = @cerrado_4,
								dpiespag_4 = '',
								atencion_4 = '',
								detraccion_4 = 0, 
								coddetrac_4 = '',
								Aprobado_4 = 0,
								userapro_4 = '',
								Tipo_4 = '01'
			WHERE CCIA_4 = @CCIA_4 AND NCOM_4 = @ncom_4 ", conx)
        cmd15.Parameters.AddWithValue("@CCIA_4", Trim(Label16.Text))
        cmd15.Parameters.AddWithValue("@ncom_4", Trim(TextBox1.Text) & Trim(TextBox2.Text))
        cmd15.Parameters.AddWithValue("@fich_4", Trim(TextBox3.Text))
        cmd15.Parameters.AddWithValue("@fcom_4", DateTimePicker1.Value.Date)
        cmd15.Parameters.AddWithValue("@fcome_4", DateTimePicker2.Value.Date)
        cmd15.Parameters.AddWithValue("@cond_4", ComboBox1.SelectedValue.ToString)
        cmd15.Parameters.AddWithValue("@tcam_4", Convert.ToDouble(TextBox9.Text))
        If ComboBox2.Text = "SOLES" Then
            cmd15.Parameters.AddWithValue("@mone_4", 1)
        Else
            cmd15.Parameters.AddWithValue("@mone_4", 2)
        End If
        If ComboBox2.Text = "SOLES" Then
            cmd15.Parameters.AddWithValue("@tot1_4", Convert.ToDouble(TextBox12.Text))
            cmd15.Parameters.AddWithValue("@tot2_4", Convert.ToDouble(TextBox12.Text) / Convert.ToDouble(TextBox9.Text))
            cmd15.Parameters.AddWithValue("@vvta1_4", Convert.ToDouble(TextBox10.Text))
            cmd15.Parameters.AddWithValue("@vvta2_4", Convert.ToDouble(TextBox10.Text) / Convert.ToDouble(TextBox9.Text))
            cmd15.Parameters.AddWithValue("@igv1_4", Convert.ToDouble(TextBox11.Text))
            cmd15.Parameters.AddWithValue("@igv2_4", Convert.ToDouble(TextBox11.Text) / Convert.ToDouble(TextBox9.Text))
        Else
            cmd15.Parameters.AddWithValue("@tot1_4", Convert.ToDouble(TextBox12.Text) * Convert.ToDouble(TextBox9.Text))
            cmd15.Parameters.AddWithValue("@tot2_4", Convert.ToDouble(TextBox12.Text))
            cmd15.Parameters.AddWithValue("@vvta1_4", Convert.ToDouble(TextBox10.Text) * Convert.ToDouble(TextBox9.Text))
            cmd15.Parameters.AddWithValue("@vvta2_4", Convert.ToDouble(TextBox10.Text))
            cmd15.Parameters.AddWithValue("@igv1_4", Convert.ToDouble(TextBox11.Text) * Convert.ToDouble(TextBox9.Text))
            cmd15.Parameters.AddWithValue("@igv2_4", Convert.ToDouble(TextBox11.Text))
        End If
        If CheckBox3.Checked = True Then
            cmd15.Parameters.AddWithValue("@cigv_4", 1)
        Else
            cmd15.Parameters.AddWithValue("@cigv_4", 0)
        End If

        cmd15.Parameters.AddWithValue("@gloa_4", Trim(TextBox7.Text))
        If CheckBox2.Checked = True Then
            cmd15.Parameters.AddWithValue("@aigv_4", 1)
        Else
            cmd15.Parameters.AddWithValue("@aigv_4", 0)
        End If

        cmd15.Parameters.AddWithValue("@lugar_4", Trim(TextBox8.Text))
        cmd15.Parameters.AddWithValue("@subfase_4", Trim(TextBox14.Text))
        cmd15.Parameters.AddWithValue("@flag_4", 1)
        cmd15.Parameters.AddWithValue("@cerrado_4", 0)
        cmd15.ExecuteNonQuery()
        abrir()
        Dim agregar12 As String = "DELETE FROM custom_vianny.dbo.lap0400 Where CCIA_4a = '" + Trim(Label16.Text) + "' And ncom_4a= '" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "'"
        ELIMINAR(agregar12)
        Dim u As Integer
        u = DataGridView1.Rows.Count

        For i = 0 To u - 1

            Dim cmd16 As New SqlCommand("INSERT INTO custom_vianny.dbo.lap0400 ( ccia_4a , ncom_4a , item_4a,cant_4a  , linea_4a , linea2_4a,preun1_4a , tot1_4a , preun2_4a ,tot2_4a , nombl_4a , ccos_4a , obser1_4a , obser2_4a , obser3_4a , obser4_4a , obser5_4a,pedido_4a , flag_4a , program_4a , pieza_4a , undm_4a , ocorte_4a , tiptar_4a,ccosto_4a,VVTA,IGV,TOTAL )
							                                            VALUES( @CCIA_4 , @ncom_4 ,@item_4a,@cant_4a ,@linea_4a , ''       ,@preun1_4a, @tot1_4 ,@preun2_4a ,@tot2_4 , ''       , ''      ,@obser1_4a, @obser2_4a , @obser3_4a, @obser4_4a,@obser5_4a,@pedido_4a,@flag_4a  ,@program_4a, ''       ,@undm_4a , ''        , ''       ,''       ,@VVTA,@IGV,@TOTAL )", conx)
            cmd16.Parameters.AddWithValue("@CCIA_4", Trim(Label16.Text))
            cmd16.Parameters.AddWithValue("@ncom_4", Trim(TextBox1.Text) & Trim(TextBox2.Text))
            cmd16.Parameters.AddWithValue("@item_4a", Trim(DataGridView1.Rows(i).Cells(0).Value))
            cmd16.Parameters.AddWithValue("@cant_4a", Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
            cmd16.Parameters.AddWithValue("@linea_4a", Trim(DataGridView1.Rows(i).Cells(4).Value))
            If ComboBox2.Text = "SOLES" Then
                cmd16.Parameters.AddWithValue("@preun1_4a", Convert.ToDouble(DataGridView1.Rows(i).Cells(8).Value))
                cmd16.Parameters.AddWithValue("@tot1_4", Convert.ToDouble(DataGridView1.Rows(i).Cells(9).Value))
                cmd16.Parameters.AddWithValue("@preun2_4a", Convert.ToDouble(DataGridView1.Rows(i).Cells(8).Value) / Convert.ToDouble(TextBox9.Text))
                cmd16.Parameters.AddWithValue("@tot2_4", Convert.ToDouble(DataGridView1.Rows(i).Cells(9).Value) / Convert.ToDouble(TextBox9.Text))
            Else
                cmd16.Parameters.AddWithValue("@preun1_4a", Convert.ToDouble(DataGridView1.Rows(i).Cells(8).Value) * Convert.ToDouble(TextBox9.Text))
                cmd16.Parameters.AddWithValue("@tot1_4", Convert.ToDouble(DataGridView1.Rows(i).Cells(9).Value) * Convert.ToDouble(TextBox9.Text))
                cmd16.Parameters.AddWithValue("@preun2_4a", Convert.ToDouble(DataGridView1.Rows(i).Cells(8).Value))
                cmd16.Parameters.AddWithValue("@tot2_4", Convert.ToDouble(DataGridView1.Rows(i).Cells(9).Value))
            End If



            cmd16.Parameters.AddWithValue("@pedido_4a", Trim(DataGridView1.Rows(i).Cells(2).Value))
            cmd16.Parameters.AddWithValue("@flag_4a", 1)
            cmd16.Parameters.AddWithValue("@program_4a", Trim(DataGridView1.Rows(i).Cells(1).Value))
            cmd16.Parameters.AddWithValue("@undm_4a", Trim(DataGridView1.Rows(i).Cells(6).Value))
            cmd16.Parameters.AddWithValue("@VVTA", DataGridView1.Rows(i).Cells(12).Value)
            cmd16.Parameters.AddWithValue("@IGV", DataGridView1.Rows(i).Cells(13).Value)
            cmd16.Parameters.AddWithValue("@TOTAL", DataGridView1.Rows(i).Cells(14).Value)
            cmd16.Parameters.AddWithValue("@obser1_4a", Trim(DataGridView1.Rows(i).Cells(11).Value))
            cmd16.Parameters.AddWithValue("@obser2_4a", Trim(DataGridView1.Rows(i).Cells(15).Value))
            cmd16.Parameters.AddWithValue("@obser3_4a", Trim(DataGridView1.Rows(i).Cells(16).Value))
            cmd16.Parameters.AddWithValue("@obser4_4a", Trim(DataGridView1.Rows(i).Cells(17).Value))
            cmd16.Parameters.AddWithValue("@obser5_4a", Trim(DataGridView1.Rows(i).Cells(18).Value))
            cmd16.ExecuteNonQuery()
        Next
        MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
        Dim hj2 As New flog
        Dim hj1 As New vlog
        hj1.gmodulo = "ORDEN_SERVICIO"
        hj1.gcuser = MDIParent1.Label3.Text
        If Label22.Text = "0" Then
            hj1.gaccion = 1
        Else
            hj1.gaccion = 2
        End If

        hj1.gpc = My.Computer.Name
        hj1.gfecha = DateTimePicker4.Value
        hj1.gdetalle = TextBox1.Text & TextBox2.Text
        hj1.gccia = Label16.Text
        hj2.insertar_log(hj1)
        MessageBox.Show("Datos registrados correctamente ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA IMPRIMIR LA ORDEN DE SERVICIO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then

            FORM_OS.TextBox1.Text = TextBox1.Text & TextBox2.Text
            FORM_OS.TextBox2.Text = TextBox1.Text & TextBox2.Text
            FORM_OS.TextBox3.Text = Label16.Text
            FORM_OS.TextBox4.Text = "\\172.21.0.1\erpcaesoft$\LIB.APPSV2\imagenes\"

            FORM_OS.Show()
        End If
        TextBox1.Text = ComboBox3.SelectedValue.ToString
        abrir()
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAlluLTIMONumeroOrdendeServicio '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        Rsr1991.Read()
        TextBox2.Text = Rsr1991(0)
        Rsr1991.Close()
        TextBox2.Select()
        Select Case ComboBox3.SelectedValue.ToString
            Case "C3" : TextBox14.Text = "0301"
            Case "D4" : TextBox14.Text = "0421"
            Case "E5" : TextBox14.Text = "0522"
        End Select
        Select Case ComboBox3.SelectedValue.ToString
            Case "C3" : TextBox15.Text = "CORTE"
            Case "D4" : TextBox15.Text = "CONFECCION"
            Case "E5" : TextBox15.Text = "ACABADO"
        End Select
        LIMPIAR()
        Label21.Visible = False
        DateTimePicker1.Value = DateTimePicker4.Value
        DateTimePicker2.Value = DateTimePicker4.Value
        CheckBox1.Checked = False
    End Sub
    Sub LIMPIAR()
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        DataGridView1.Rows.Clear()
        TextBox7.Text = ""

        TextBox8.Text = ""
    End Sub
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        'Dim corre As String
        'jj.gvendedor = TextBox8.Text
        'corre = fk.buscar_correo(jj)

        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA ORDEN DE SERVICIO</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> ANULADO <br/>
<FONT COLOR='blue'>* AREA : </FONT>" + Trim(ComboBox3.Text) + " <br/>
<FONT COLOR='blue'>* ORDEN : </FONT>" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + " <br/>
<FONT COLOR='blue'>* USUARIO : </FONT>" + Trim(MDIParent1.Label3.Text) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(MDIParent1.Label7.Text) + " <br/>
<FONT COLOR='blue'>* TALLER : </FONT>" + Trim(TextBox4.Text) + "<br/>
</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("fmestanza@viannysac.com,hrivera@viannysac.com")
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "ORDEN SERVICIO N°" & Trim(TextBox1.Text) & Trim(TextBox2.Text)
            .Priority = System.Net.Mail.MailPriority.Normal
        End With

        With smtp
            '.EnableSsl = True
            .Port = "25"
            .Host = "mail.onehostingperu.com"
            .Credentials = New Net.NetworkCredential("admin@viannysac.com", "I9!?@ni2;+go")
            .Send(message)
        End With

        Try
            MessageBox.Show("Su mensaje de correo ha sido enviado.", "Correo enviado",
                             MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error al enviar correo",
                             MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If CheckBox1.Checked = True Then
            MsgBox("NO SE PUEDE ANULAR LA ORDEN DE SERVICIO QUE ESTA APROBADA")
        Else
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("QUIERES ANULAR NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Dim cmd1577 As New SqlCommand("update custom_vianny.dbo.LAG0400 set flag_4=0 where ncom_4 =@ncom4 and ccia_4=@ccia_4", conx)
                cmd1577.Parameters.AddWithValue("@ncom4", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd1577.Parameters.AddWithValue("@ccia_4", Trim(Label16.Text))
                cmd1577.ExecuteNonQuery()

                Dim cmd15776 As New SqlCommand("update custom_vianny.dbo.lap0400 set flag_4a=0 where ncom_4a =@ncom4 and ccia_4a=@ccia_4", conx)
                cmd15776.Parameters.AddWithValue("@ncom4", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd15776.Parameters.AddWithValue("@ccia_4", Trim(Label16.Text))
                cmd15776.ExecuteNonQuery()
                MsgBox("SE ANULO LA INFORMACION CORRECTAMENTE")
                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "ORDEN_SERVICIO"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = 3
                hj1.gpc = My.Computer.Name
                hj1.gfecha = DateTimePicker3.Value
                hj1.gdetalle = TextBox1.Text & TextBox2.Text
                hj1.gccia = Label16.Text
                hj2.insertar_log(hj1)
                ENVIAR_CORREO()
                TextBox1.Text = ComboBox3.SelectedValue.ToString
                abrir()
                Dim Rsr1991 As SqlDataReader
                Dim sql1011 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAlluLTIMONumeroOrdendeServicio '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "'"
                Dim cmd1011 As New SqlCommand(sql1011, conx)
                Rsr1991 = cmd1011.ExecuteReader()
                Rsr1991.Read()
                TextBox2.Text = Rsr1991(0)
                Rsr1991.Close()
                TextBox2.Select()
                Select Case ComboBox3.SelectedValue.ToString
                    Case "C3" : TextBox14.Text = "0301"
                    Case "D4" : TextBox14.Text = "0421"
                    Case "E5" : TextBox14.Text = "0522"
                End Select
                Select Case ComboBox3.SelectedValue.ToString
                    Case "C3" : TextBox15.Text = "CORTE"
                    Case "D4" : TextBox15.Text = "CONFECCION"
                    Case "E5" : TextBox15.Text = "ACABADO"
                End Select
                LIMPIAR()
            End If
        End If



    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        LIMPIAR()
        TextBox1.Text = ComboBox3.SelectedValue.ToString
        DateTimePicker1.Value = DateTimePicker4.Value
        DateTimePicker2.Value = DateTimePicker4.Value
        abrir()
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAlluLTIMONumeroOrdendeServicio '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        Rsr1991.Read()
        TextBox2.Text = Rsr1991(0)
        Rsr1991.Close()
        TextBox2.Select()
        Select Case ComboBox3.SelectedValue.ToString
            Case "C3" : TextBox14.Text = "0301"
            Case "D4" : TextBox14.Text = "0421"
            Case "E5" : TextBox14.Text = "0522"
        End Select
        Select Case ComboBox3.SelectedValue.ToString
            Case "C3" : TextBox15.Text = "CORTE"
            Case "D4" : TextBox15.Text = "CONFECCION"
            Case "E5" : TextBox15.Text = "ACABADO"
        End Select
        Label21.Visible = False
        CheckBox1.Checked = False
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

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

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim func As New ftcambio
        Dim dts As New vtcambio

        dts.gfecha = DateTimePicker1.Text
        TextBox9.Text = Func.mostrar_tipo_cambio(dts)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        log.Label1.Text = Trim(TextBox1.Text) & Trim(TextBox2.Text)
        log.Label2.Text = "ORDEN_SERVICIO"
        log.Label3.Text = Label16.Text
        log.Show()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "C" Then
            pproductos.CCIA.Text = Label16.Text
            pproductos.ALMACEN.Text = 10
            pproductos.ANO.Text = Label11.Text
            pproductos.FECHA.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = Me.ComboBox3.Text
            pproductos.Label3.Text = 3
            pproductos.Label5.Text = e.RowIndex
            pproductos.Show()
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Obs" Then
            OB_SC.Label1.Text = e.RowIndex
            OB_SC.Label2.Text = 1
            OB_SC.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value)
            OB_SC.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(15).Value)
            OB_SC.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(16).Value)
            OB_SC.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(17).Value)
            OB_SC.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(18).Value)
            OB_SC.ShowDialog()
        End If
    End Sub

    Private Sub TextBox2_Invalidated(sender As Object, e As InvalidateEventArgs) Handles TextBox2.Invalidated

    End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.F1 Then
            Procesos.Label2.Text = Trim(Label16.Text)
            Procesos.Label1.Text = "5"
            Procesos.Show()
            TextBox8.Select()
        End If
    End Sub

    Private Sub DateTimePicker1_Click(sender As Object, e As EventArgs) Handles DateTimePicker1.Click
        If Trim(ComboBox3.Text) = "C3    | CORTE" Then
            DateTimePicker2.Value = DateTimePicker1.Value.AddDays(2)
        Else
            If Trim(ComboBox3.Text) = "D4    | CONFECCION" Then
                DateTimePicker2.Value = DateTimePicker1.Value.AddDays(8)
            Else
                If Trim(ComboBox3.Text) = "E5    | ACABADO" Then
                    DateTimePicker2.Value = DateTimePicker1.Value.AddDays(2)
                Else
                    If Trim(ComboBox3.Text) = "H8    | LAVANDERIA" Then
                        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(3)
                    Else
                        DateTimePicker2.Value = DateTimePicker1.Value
                    End If
                End If
            End If
        End If
    End Sub
End Class