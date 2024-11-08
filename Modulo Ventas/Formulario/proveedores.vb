Imports System.Data.SqlClient
Public Class proveedores
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
    Dim Rsr2 As SqlDataReader

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim va, va1 As Double
        Dim hj As Integer
        hj = DataGridView1.Rows.Count
        For i = 0 To hj - 1
            va = va + DataGridView1.Rows(i).Cells(9).Value
            va1 = va1 + DataGridView1.Rows(i).Cells(10).Value

        Next
        TextBox3.Text = va
        TextBox4.Text = va1
    End Sub

    Private Sub proveedores_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "PROVEEDOR" & " like '%" & TextBox5.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                da = dv.Table(0).Table
                'DataGridView1.Columns(12).Visible = False
                'DataGridView1.Columns(13).Visible = False
                'DataGridView1.Columns(1).Frozen = True
                'DataGridView1.Columns(2).Frozen = True
                'DataGridView1.Columns(3).Frozen = True
                'DataGridView1.Columns(4).Frozen = True
                'DataGridView1.Columns(5).Frozen = True
                'DataGridView1.Columns(6).Frozen = True
                'DataGridView1.Columns(7).Frozen = True
                'DataGridView1.Columns(8).Frozen = True
                'DataGridView1.Columns(9).Frozen = True
                'DataGridView1.Columns(1).Width = 70
                'DataGridView1.Columns(4).Width = 80
                'DataGridView1.Columns(5).Width = 80
                'DataGridView1.Columns(6).Width = 80
                'DataGridView1.Columns(7).Width = 80
                'DataGridView1.Columns(8).Width = 50
                'DataGridView1.Columns(9).Width = 45
                'DataGridView1.Columns(1).ReadOnly = True
                'DataGridView1.Columns(2).ReadOnly = True
                'DataGridView1.Columns(3).ReadOnly = True
                'DataGridView1.Columns(4).ReadOnly = True
                'DataGridView1.Columns(5).ReadOnly = True
                'DataGridView1.Columns(6).ReadOnly = True
                'DataGridView1.Columns(7).ReadOnly = True
                'DataGridView1.Columns(8).ReadOnly = True
                'DataGridView1.Columns(9).ReadOnly = True
                Dim u As Integer

                For i = 0 To u - 1
                    If Trim(DataGridView1.Rows(i).Cells("SE").Value) = 1 Then
                        DataGridView1.Rows(i).Cells(0).Value = True
                    End If
                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar_cuenta()
        Try
            Dim ds As New DataSet
            Dim K As Integer


            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "CUENTA" & " like '%" & TextBox6.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(6).Frozen = True
                DataGridView1.Columns(7).Frozen = True
                DataGridView1.Columns(8).Frozen = True
                DataGridView1.Columns(9).Frozen = True
                DataGridView1.Columns(1).Width = 70
                DataGridView1.Columns(4).Width = 80
                DataGridView1.Columns(5).Width = 80
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(8).Width = 50
                DataGridView1.Columns(9).Width = 45
                DataGridView1.Columns(1).ReadOnly = True
                DataGridView1.Columns(2).ReadOnly = True
                DataGridView1.Columns(3).ReadOnly = True
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
                DataGridView1.Columns(7).ReadOnly = True
                DataGridView1.Columns(8).ReadOnly = True
                DataGridView1.Columns(9).ReadOnly = True
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs)
        'buscar()
        bs.Filter = String.Format("PROVEEDOR LIKE '%{0}%'", TextBox5.Text)
        Dim semanaa, smpea, smpe, semana As Integer
        semanaa = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        semana = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        smpea = 53 - semanaa
        smpe = 53 - semana
        Dim pl As Integer
        pl = DataGridView2.Rows.Count
        Dim pl1, su, semar, semar1 As Integer
        Dim can, can1 As Double
        Dim ano, ano1 As Integer
        Dim va, va1, va2, va3 As Double


        su = DataGridView1.Rows.Count
        For i1 = 0 To su - 1
            If Trim(DataGridView1.Rows(i1).Cells(6).Value) = "" Then
                semar = 1
                ano = "0"
            Else
                semar = DatePart("ww", DataGridView1.Rows(i1).Cells(6).Value, vbMonday, vbFirstFourDays)
                ano = Year(DataGridView1.Rows(i1).Cells(6).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(7).Value) = "" Then
                semar1 = 1
                ano1 = "0"
            Else
                semar1 = DatePart("ww", DataGridView1.Rows(i1).Cells(7).Value, vbMonday, vbFirstFourDays)
                ano1 = Year(DataGridView1.Rows(i1).Cells(7).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(7).Value) = "" Then
                If semar < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano < Year(DateTimePicker1.Value) Then
                    DataGridView1.Rows(i1).Cells(11).Value = DataGridView1.Rows(i1).Cells(10).Value
                Else
                    DataGridView1.Rows(i1).Cells(11).Value = 0.0
                    For i2 = 16 To smpe + 15

                        If DataGridView1.Columns(i2).HeaderText = "SEM " & semar Then
                            DataGridView1.Rows(i1).Cells(i2).Value = DataGridView1.Rows(i1).Cells(10).Value
                            Exit For
                        End If
                    Next
                End If
            Else

                If semar1 < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano1 < Year(DateTimePicker1.Value) Then
                    DataGridView1.Rows(i1).Cells(11).Value = DataGridView1.Rows(i1).Cells(10).Value
                Else
                    DataGridView1.Rows(i1).Cells(11).Value = 0.0
                    For i2 = 16 To smpe + 15

                        If DataGridView1.Columns(i2).HeaderText = "SEM " & semar1 Then
                            DataGridView1.Rows(i1).Cells(i2).Value = DataGridView1.Rows(i1).Cells(10).Value
                            Exit For
                        End If
                    Next
                End If
            End If



            If Trim(DataGridView1.Rows(i1).Cells(9).Value) = "S/." Then
                va = va + Val(DataGridView1.Rows(i1).Cells(10).Value)
            Else
                va2 = va2 + Val(DataGridView1.Rows(i1).Cells(10).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(9).Value) = "S/." Then
                va1 = va1 + Val(DataGridView1.Rows(i1).Cells(11).Value)
            Else
                va3 = va3 + Val(DataGridView1.Rows(i1).Cells(11).Value)
            End If

            'If Trim(DataGridView1.Rows(i1).Cells(7).Value).Length > 0 And Trim(DataGridView1.Rows(i1).Cells(14).Value) = "" Then
            '    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
            '    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
            'End If

            If Trim(DataGridView1.Rows(i1).Cells(7).Value).Length > 0 And Trim(DataGridView1.Rows(i1).Cells(14).Value) = "" Then
                If Trim(DataGridView1.Rows(i1).Cells(7).Value) < DateTimePicker1.Value Then
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                Else
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.DarkCyan
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                End If
            Else
                If Trim(DataGridView1.Rows(i1).Cells(15).Value) = "" Then
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.DarkGray
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                Else

                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.White
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.Black
                End If
            End If
            If Trim(DataGridView1.Rows(i1).Cells(14).Value) = "1" Then
                DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Yellow
                DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.Black
                DataGridView1.Rows(i1).Cells(0).Value = True
            End If
            If Trim(DataGridView1.Rows(i1).Cells(17).Value) = "1" Then
                DataGridView1.Rows(i1).Visible = False
            End If
        Next
        For p1 = 1 To smpea

            For p2 = 16 To smpea + 15

                If Trim(DataGridView2.Columns(p1).HeaderText) = Trim(DataGridView1.Columns(p2).HeaderText) Then
                    pl1 = DataGridView1.Rows.Count
                    For p3 = 0 To pl1 - 1
                        If Trim(DataGridView1.Rows(p3).Cells(9).Value) = "S/." Then
                            can = can + Val(DataGridView1.Rows(p3).Cells(p2).Value)
                        Else
                            can1 = can1 + Val(DataGridView1.Rows(p3).Cells(p2).Value)
                        End If
                    Next
                    DataGridView2.Rows(0).Cells(p1).Value = can
                    DataGridView2.Rows(1).Cells(p1).Value = can1
                Else
                    can = 0
                    can1 = 0
                End If
            Next

        Next
        TextBox3.Text = va.ToString("N2")
        TextBox4.Text = va1.ToString("N2")
        TextBox7.Text = va2.ToString("N2")
        TextBox8.Text = va3.ToString("N2")
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Dim bs As New BindingSource()

    Dim da As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = ""
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("exec cuentasxpagar '" + Trim(TextBox1.Text) + "'", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs

            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(14).Visible = False
            DataGridView1.Columns(15).Visible = False
            'DataGridView1.Columns(17).Visible = False
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            DataGridView1.Columns(4).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            DataGridView1.Columns(6).Frozen = True
            DataGridView1.Columns(7).Frozen = True
            DataGridView1.Columns(8).Frozen = True
            DataGridView1.Columns(9).Frozen = True
            DataGridView1.Columns(1).Width = 70
            DataGridView1.Columns(4).Width = 80
            DataGridView1.Columns(5).Width = 80
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(8).Width = 50
            DataGridView1.Columns(9).Width = 45
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).ReadOnly = True
            DataGridView1.Columns(4).ReadOnly = True
            DataGridView1.Columns(5).ReadOnly = True
            DataGridView1.Columns(6).ReadOnly = True
            DataGridView1.Columns(7).ReadOnly = True
            DataGridView1.Columns(8).ReadOnly = True
            DataGridView1.Columns(9).ReadOnly = True

        Else
            DataGridView1.DataSource = Nothing
        End If

        'Dim sql102 As String = "exec cuentasxpagar '" + Trim(TextBox1.Text) + "'"
        'Dim cmd102 As New SqlCommand(sql102, conx)
        'cmd102.CommandTimeout = 200
        'Rsr2 = cmd102.ExecuteReader()
        Dim i5, semana, smpe, su, semar, semana1, semar1 As Integer
        'i5 = 0



        'While Rsr2.Read()
        '    DataGridView1.Rows.Add()

        '    DataGridView1.Rows(i5).Cells(1).Value = Rsr2(0)
        '    DataGridView1.Rows(i5).Cells(2).Value = Rsr2(2)
        '    DataGridView1.Rows(i5).Cells(3).Value = Rsr2(3)
        '    DataGridView1.Rows(i5).Cells(4).Value = Rsr2(11)
        '    DataGridView1.Rows(i5).Cells(5).Value = Rsr2(12)
        '    DataGridView1.Rows(i5).Cells(6).Value = Rsr2(13)

        '    DataGridView1.Rows(i5).Cells(7).Value = Rsr2(30)
        '    DataGridView1.Rows(i5).Cells(8).Value = Rsr2(14)
        '    DataGridView1.Rows(i5).Cells(9).Value = Rsr2(15)
        '    DataGridView1.Rows(i5).Cells(12).Value = Rsr2(10)
        '    DataGridView1.Rows(i5).Cells(13).Value = Rsr2(9)
        '    If Trim(DataGridView1.Rows(i5).Cells(9).Value) = "S/." Then
        '        DataGridView1.Rows(i5).Cells(10).Value = Rsr2(19)
        '    Else

        '        DataGridView1.Rows(i5).Cells(10).Value = Rsr2(20)


        '    End If

        '    i5 = i5 + 1


        '    'DataGridView2.Rows(i5).Cells(6).Value = (Rsr2(12) * DataGridView1.Rows(0).Cells(10).Value) / 100
        'End While
        'Table.Load(Rsr2)
        'Rsr2.Close()

        semana = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        semana1 = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        smpe = 53 - semana
        For i = 14 To smpe + 13
            DataGridView1.Columns.Add("SEM " & semana, "SEM " & semana)

            semana = semana + 1
        Next

        For i8 = 1 To smpe
            DataGridView2.Columns.Add("SEM " & semana1, "SEM " & semana1)

            semana1 = semana1 + 1
        Next
        DataGridView2.Rows.Add(2)
        DataGridView2.Rows(0).Cells(0).Value = "S/."
        DataGridView2.Rows(1).Cells(0).Value = "$."
        su = DataGridView1.Rows.Count
        Dim ano, ano1 As Integer
        Dim va, va1, va2, va3 As Double
        For i1 = 0 To su - 1
            If Trim(DataGridView1.Rows(i1).Cells(6).Value) = "" Then
                semar = 1
                ano = "0"
            Else
                semar = DatePart("ww", DataGridView1.Rows(i1).Cells(6).Value, vbMonday, vbFirstFourDays)
                ano = Year(DataGridView1.Rows(i1).Cells(6).Value)
            End If

            If Trim(DataGridView1.Rows(i1).Cells(7).Value) = "" Then
                semar1 = 1
                ano1 = "0"
            Else
                semar1 = DatePart("ww", DataGridView1.Rows(i1).Cells(7).Value, vbMonday, vbFirstFourDays)
                ano1 = Year(DataGridView1.Rows(i1).Cells(7).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(7).Value) = "" Then
                If Trim(DataGridView1.Rows(i1).Cells(15).Value) = "1" Then
                    If semar < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano < Year(DateTimePicker1.Value) Then
                        DataGridView1.Rows(i1).Cells(11).Value = DataGridView1.Rows(i1).Cells(10).Value
                    Else
                        DataGridView1.Rows(i1).Cells(11).Value = 0.0
                        For i2 = 16 To smpe + 15

                            If DataGridView1.Columns(i2).HeaderText = "SEM " & semar Then
                                DataGridView1.Rows(i1).Cells(i2).Value = DataGridView1.Rows(i1).Cells(10).Value
                                Exit For
                            End If
                        Next
                    End If
                Else

                End If

            Else
                If Trim(DataGridView1.Rows(i1).Cells(15).Value) = "1" Then
                    If semar1 < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano1 < Year(DateTimePicker1.Value) Then
                        DataGridView1.Rows(i1).Cells(11).Value = DataGridView1.Rows(i1).Cells(10).Value
                    Else
                        DataGridView1.Rows(i1).Cells(11).Value = 0.0
                        For i2 = 16 To smpe + 15

                            If DataGridView1.Columns(i2).HeaderText = "SEM " & semar1 Then
                                DataGridView1.Rows(i1).Cells(i2).Value = DataGridView1.Rows(i1).Cells(10).Value
                                Exit For
                            End If
                        Next
                    End If
                Else

                End If

            End If


            If DataGridView1.Rows(i1).Cells(10).Value Is DBNull.Value Then
                DataGridView1.Rows(i1).Cells(10).Value = 0
            End If
            If DataGridView1.Rows(i1).Cells(11).Value Is DBNull.Value Then
                DataGridView1.Rows(i1).Cells(11).Value = 0
            End If
            If Trim(DataGridView1.Rows(i1).Cells(9).Value) = "S/." Then
                va = va + Val(DataGridView1.Rows(i1).Cells(10).Value)
            Else
                va2 = va2 + Val(DataGridView1.Rows(i1).Cells(10).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(9).Value) = "S/." Then
                va1 = va1 + Val(DataGridView1.Rows(i1).Cells(11).Value)
            Else
                va3 = va3 + Val(DataGridView1.Rows(i1).Cells(11).Value)
            End If



        Next
        TextBox3.Text = va.ToString("N2")
        TextBox4.Text = va1.ToString("N2")
        TextBox7.Text = va2.ToString("N2")
        TextBox8.Text = va3.ToString("N2")
        Dim semanaa, smpea As Integer
        semanaa = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        smpea = 53 - semanaa
        Dim pl As Integer
        pl = DataGridView2.Rows.Count
        Dim pl1 As Integer
        Dim can, can1 As Double
        pl1 = DataGridView1.Rows.Count

        For p1 = 1 To smpea

            For p2 = 17 To smpea + 15

                If Trim(DataGridView2.Columns(p1).HeaderText) = Trim(DataGridView1.Columns(p2).HeaderText) Then

                    For p3 = 0 To pl1 - 1
                        If Trim(DataGridView1.Rows(p3).Cells(9).Value) = "S/." Then
                            can = can + Val(DataGridView1.Rows(p3).Cells(p2).Value)
                        Else
                            can1 = can1 + Val(DataGridView1.Rows(p3).Cells(p2).Value)
                        End If
                    Next
                    DataGridView2.Rows(0).Cells(p1).Value = can
                    DataGridView2.Rows(1).Cells(p1).Value = can1
                Else
                    can = 0
                    can1 = 0
                End If
            Next

        Next
        Dim Lf As Integer
        Lf = DataGridView1.Rows.Count
        For k = 0 To Lf - 1

            If Trim(DataGridView1.Rows(k).Cells(7).Value).Length > 0 Then
                If Trim(DataGridView1.Rows(k).Cells(7).Value) < DateTimePicker1.Value Then
                    DataGridView1.Rows(k).DefaultCellStyle.BackColor = Color.Red
                    DataGridView1.Rows(k).DefaultCellStyle.ForeColor = Color.White
                Else
                    DataGridView1.Rows(k).DefaultCellStyle.BackColor = Color.DarkCyan
                    DataGridView1.Rows(k).DefaultCellStyle.ForeColor = Color.White
                End If
            Else
                If Trim(DataGridView1.Rows(k).Cells(15).Value) = "" Then
                    DataGridView1.Rows(k).DefaultCellStyle.BackColor = Color.DarkGray
                    DataGridView1.Rows(k).DefaultCellStyle.ForeColor = Color.White
                End If

            End If
            If Trim(DataGridView1.Rows(k).Cells(14).Value) = "1" Then
                DataGridView1.Rows(k).DefaultCellStyle.BackColor = Color.Yellow
                DataGridView1.Rows(k).DefaultCellStyle.ForeColor = Color.Black
                DataGridView1.Rows(k).Cells(0).Value = True
            End If
            If Trim(DataGridView1.Rows(k).Cells(17).Value) = "1" Then
                DataGridView1.Rows(k).Visible = False
            End If
        Next
        Button1.Enabled = False
        Button4.Enabled = True
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex = -1 Then
            MsgBox("Esta Funcion esta Bloqueada")
            For Each col As DataGridViewColumn In DataGridView1.Columns

                col.SortMode = DataGridViewColumnSortMode.NotSortable

            Next
        Else
            If DataGridView1.Columns(e.ColumnIndex).HeaderText = "F_PROGRA" Then
                If DataGridView1.Rows(e.RowIndex).Cells(0).Value = True Then
                    Dim num1, num2 As Integer
                    num1 = e.RowIndex.ToString
                    num2 = e.ColumnIndex.ToString

                    F_PROGRAMACION.Label3.Text = num1
                    F_PROGRAMACION.Label2.Text = 1
                    F_PROGRAMACION.Label4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(9).Value)
                    F_PROGRAMACION.ShowDialog()
                End If
            Else
                If DataGridView1.Columns(e.ColumnIndex).HeaderText = "DOCUM" Then
                    Dim respuesta As DialogResult
                    Dim cco As String
                    respuesta = MessageBox.Show("DESEA VALIDAR COMO RECIBIDO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        'MsgBox(Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value))
                        'MsgBox(Trim(Label1.Text))
                        'MsgBox(Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value))
                        'MsgBox(Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value))
                        'MsgBox(Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value))
                        'MsgBox(Trim(TextBox1.Text))
                        Dim cmd17 As New SqlCommand("update custom_vianny.dbo.cac3p00 set fuen_3a=@fuen_3a where ncom_3a =@ncom_3 and ccia_3a =@ccia_3 and cperiodo_3a =@cperiodo_3 and fich_3a =@ruc AND ccom_3a =@ccom_3 and ndoc_3a =@ndoc_3a and cuen_3a =@cuenta", conx)
                        cmd17.Parameters.AddWithValue("@fuen_3a", "1")
                        cmd17.Parameters.AddWithValue("@ncom_3", Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value))
                        cmd17.Parameters.AddWithValue("@ccia_3", Trim(Label1.Text))
                        cmd17.Parameters.AddWithValue("@cperiodo_3", Trim(TextBox1.Text))
                        cmd17.Parameters.AddWithValue("@ruc", Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value))
                        cmd17.Parameters.AddWithValue("@ndoc_3a", Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value))
                        cmd17.Parameters.AddWithValue("@cuenta", Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value))

                        Dim Rsr1991 As SqlDataReader
                        Dim sql1011 As String = "select ncom_3i from custom_vianny.dbo.cag3i00 where ccom_3i ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value) + "' and ccia_3i='" + Label1.Text + "'"
                        Dim cmd1011 As New SqlCommand(sql1011, conx)
                        Rsr1991 = cmd1011.ExecuteReader()
                        Rsr1991.Read()
                        cco = Rsr1991(0)
                        Rsr1991.Close()
                        cmd17.Parameters.AddWithValue("@ccom_3", Trim(cco))
                        'MsgBox(Trim(cco))
                        cmd17.ExecuteNonQuery()
                        DataGridView1.Rows(e.RowIndex).Cells("fuen_3a").Value = 1
                        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
                        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Black
                        MsgBox("SE REGISTRO LO SOLICITADO")
                        If cco = "500" Then
                            Dim Rsr1992 As SqlDataReader
                            Dim hj As Integer
                            Dim sql1012 As String = "select count(ctab) from tab0100 where ccia ='" + Trim(Label1.Text) + "' and ctab ='" + Trim(TextBox1.Text) + "' and cele = '" + Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value) + "' and dele ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) + "' and nele ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value) + "' and codigo ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) + "' and valo = '" + cco + "'"
                            Dim cmd1012 As New SqlCommand(sql1012, conx)
                            Rsr1992 = cmd1011.ExecuteReader()
                            Rsr1992.Read()
                            hj = Rsr1992(0)
                            Rsr1992.Close()

                            If hj = 0 Then
                                Dim cmd18 As New SqlCommand("insert into tab0100 (ccia,ctab,cele,dele,nele,codigo,valo) values(@ccia,@ctab,@cele,@dele,@nele,@codigo,@valo)", conx)
                                cmd18.Parameters.AddWithValue("@ccia", Trim(Label1.Text))
                                cmd18.Parameters.AddWithValue("@ctab", Trim(TextBox1.Text))
                                cmd18.Parameters.AddWithValue("@cele", Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value))
                                cmd18.Parameters.AddWithValue("@dele", Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value))
                                cmd18.Parameters.AddWithValue("@nele", Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value))
                                cmd18.Parameters.AddWithValue("@codigo", Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value))
                                cmd18.Parameters.AddWithValue("@valo", cco)
                                cmd18.ExecuteNonQuery()
                            Else
                                MsgBox("EL DOCUMENTO YA ESTA VALIDADO")
                            End If



                        End If

                        Dim semana, smpe, semar, semar1 As Integer
                        Dim ano, ano1 As Integer
                        semana = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)

                        smpe = 53 - semana
                        If Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value) = "" Then
                            semar = 1
                            ano = "0"
                        Else
                            semar = DatePart("ww", DataGridView1.Rows(e.RowIndex).Cells(6).Value, vbMonday, vbFirstFourDays)
                            ano = Year(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
                        End If

                        If Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value) = "" Then
                            semar1 = 1
                            ano1 = "0"
                        Else
                            semar1 = DatePart("ww", DataGridView1.Rows(e.RowIndex).Cells(7).Value, vbMonday, vbFirstFourDays)
                            ano1 = Year(DataGridView1.Rows(e.RowIndex).Cells(7).Value)
                        End If
                        If Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value) = "" Then
                            If Trim(DataGridView1.Rows(e.RowIndex).Cells(15).Value) = "1" Then
                                If semar < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano < Year(DateTimePicker1.Value) Then
                                    DataGridView1.Rows(e.RowIndex).Cells(11).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                Else
                                    DataGridView1.Rows(e.RowIndex).Cells(11).Value = 0.0
                                    For i2 = 17 To smpe + 15

                                        If DataGridView1.Columns(i2).HeaderText = "SEM " & semar Then
                                            DataGridView1.Rows(e.RowIndex).Cells(i2).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                            Exit For
                                        End If
                                    Next
                                End If
                            Else

                            End If

                        Else
                            If Trim(DataGridView1.Rows(e.RowIndex).Cells(15).Value) = "1" Then
                                If semar1 < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano1 < Year(DateTimePicker1.Value) Then
                                    DataGridView1.Rows(e.RowIndex).Cells(11).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                Else
                                    DataGridView1.Rows(e.RowIndex).Cells(11).Value = 0.0
                                    For i2 = 17 To smpe + 15

                                        If DataGridView1.Columns(i2).HeaderText = "SEM " & semar1 Then
                                            DataGridView1.Rows(e.RowIndex).Cells(i2).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                            Exit For
                                        End If
                                    Next
                                End If
                            Else

                            End If

                        End If
                    End If
                Else
                    If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PROVEEDOR" Then
                        Dim respuesta1 As DialogResult
                        Dim cco As String
                        respuesta1 = MessageBox.Show("LA OPCION SI - ELIMINAR FACTURA -- OPCION NO FACTURA EXPORTACION ", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation)
                        If (respuesta1 = Windows.Forms.DialogResult.Yes) Then
                            Dim cmd177 As New SqlCommand("update custom_vianny.dbo.cac3p00 set tefe_3a=@tefe_3a where ncom_3a =@ncom_3 and ccia_3a =@ccia_3 and cperiodo_3a =@cperiodo_3 and fich_3a =@ruc AND ccom_3a =@ccom_3 and ndoc_3a =@ndoc_3a and cuen_3a =@cuenta", conx)
                            cmd177.Parameters.AddWithValue("@tefe_3a", "1")
                            cmd177.Parameters.AddWithValue("@ncom_3", Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value))
                            cmd177.Parameters.AddWithValue("@ccia_3", Trim(Label1.Text))
                            cmd177.Parameters.AddWithValue("@cperiodo_3", Trim(TextBox1.Text))
                            cmd177.Parameters.AddWithValue("@ruc", Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value))
                            cmd177.Parameters.AddWithValue("@ndoc_3a", Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value))
                            cmd177.Parameters.AddWithValue("@cuenta", Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value))

                            Dim Rsr19911 As SqlDataReader
                            Dim sql10111 As String = "select ncom_3i from custom_vianny.dbo.cag3i00 where ccom_3i ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value) + "' and ccia_3i='" + Label1.Text + "'"
                            Dim cmd10111 As New SqlCommand(sql10111, conx)
                            Rsr19911 = cmd10111.ExecuteReader()
                            Rsr19911.Read()
                            cco = Rsr19911(0)
                            Rsr19911.Close()
                            cmd177.Parameters.AddWithValue("@ccom_3", Trim(cco))
                            'MsgBox(Trim(cco))
                            cmd177.ExecuteNonQuery()

                            Dim cmd188 As New SqlCommand("insert into facturas_eliminadas (empresa,ano,ncom,ruc,documento,cuenta,provision,sigla_voucher,observacion,estado) values (@empresa,@ano,@ncom,@ruc,@documento,@cuenta,@provision,@sigla_voucher,@observacion,@estado)", conx)
                            cmd188.Parameters.AddWithValue("@empresa", Trim(Label1.Text))
                            cmd188.Parameters.AddWithValue("@ano", Trim(TextBox1.Text))
                            cmd188.Parameters.AddWithValue("@ncom", Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value))
                            cmd188.Parameters.AddWithValue("@ruc", Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value))
                            cmd188.Parameters.AddWithValue("@documento", Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value))
                            cmd188.Parameters.AddWithValue("@cuenta", Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value))
                            cmd188.Parameters.AddWithValue("@provision", (DataGridView1.Rows(e.RowIndex).Cells(10).Value))
                            cmd188.Parameters.AddWithValue("@sigla_voucher", Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value))
                            cmd188.Parameters.AddWithValue("@observacion", "")
                            cmd188.Parameters.AddWithValue("@estado", 1)
                            cmd188.ExecuteNonQuery()
                            DataGridView1.Rows(e.RowIndex).Cells("tefe_3a").Value = 1
                            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.AliceBlue
                            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Red
                            Dim semana, smpe, semar, semar1 As Integer
                            Dim ano, ano1 As Integer
                            semana = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)

                            smpe = 53 - semana
                            If Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value) = "" Then
                                semar = 1
                                ano = "0"
                            Else
                                semar = DatePart("ww", DataGridView1.Rows(e.RowIndex).Cells(6).Value, vbMonday, vbFirstFourDays)
                                ano = Year(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
                            End If

                            If Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value) = "" Then
                                semar1 = 1
                                ano1 = "0"
                            Else
                                semar1 = DatePart("ww", DataGridView1.Rows(e.RowIndex).Cells(7).Value, vbMonday, vbFirstFourDays)
                                ano1 = Year(DataGridView1.Rows(e.RowIndex).Cells(7).Value)
                            End If
                            If Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value) = "" Then
                                If Trim(DataGridView1.Rows(e.RowIndex).Cells(15).Value) = "1" Then
                                    If semar < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano < Year(DateTimePicker1.Value) Then
                                        DataGridView1.Rows(e.RowIndex).Cells(11).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                    Else
                                        DataGridView1.Rows(e.RowIndex).Cells(11).Value = 0.0
                                        For i2 = 17 To smpe + 15

                                            If DataGridView1.Columns(i2).HeaderText = "SEM " & semar Then
                                                DataGridView1.Rows(e.RowIndex).Cells(i2).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Else

                                End If

                            Else
                                If Trim(DataGridView1.Rows(e.RowIndex).Cells(15).Value) = "1" Then
                                    If semar1 < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano1 < Year(DateTimePicker1.Value) Then
                                        DataGridView1.Rows(e.RowIndex).Cells(11).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                    Else
                                        DataGridView1.Rows(e.RowIndex).Cells(11).Value = 0.0
                                        For i2 = 17 To smpe + 15

                                            If DataGridView1.Columns(i2).HeaderText = "SEM " & semar1 Then
                                                DataGridView1.Rows(e.RowIndex).Cells(i2).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Else

                                End If

                            End If
                        Else
                            If (respuesta1 = Windows.Forms.DialogResult.No) Then


                                Dim cmd177 As New SqlCommand("update custom_vianny.dbo.cac3p00 set tefe_3a=@tefe_3a where ncom_3a =@ncom_3 and ccia_3a =@ccia_3 and cperiodo_3a =@cperiodo_3 and fich_3a =@ruc AND ccom_3a =@ccom_3 and ndoc_3a =@ndoc_3a and cuen_3a =@cuenta", conx)
                                cmd177.Parameters.AddWithValue("@tefe_3a", "1")
                                cmd177.Parameters.AddWithValue("@ncom_3", Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value))
                                cmd177.Parameters.AddWithValue("@ccia_3", Trim(Label1.Text))
                                cmd177.Parameters.AddWithValue("@cperiodo_3", Trim(TextBox1.Text))
                                cmd177.Parameters.AddWithValue("@ruc", Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value))
                                cmd177.Parameters.AddWithValue("@ndoc_3a", Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value))
                                cmd177.Parameters.AddWithValue("@cuenta", Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value))

                                Dim Rsr19911 As SqlDataReader
                                Dim sql10111 As String = "select ncom_3i from custom_vianny.dbo.cag3i00 where ccom_3i ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value) + "' and ccia_3i='" + Label1.Text + "'"
                                Dim cmd10111 As New SqlCommand(sql10111, conx)
                                Rsr19911 = cmd10111.ExecuteReader()
                                Rsr19911.Read()
                                cco = Rsr19911(0)
                                Rsr19911.Close()
                                cmd177.Parameters.AddWithValue("@ccom_3", Trim(cco))
                                'MsgBox(Trim(cco))
                                cmd177.ExecuteNonQuery()

                                Dim cmd188 As New SqlCommand("insert into facturas_eliminadas (empresa,ano,ncom,ruc,documento,cuenta,provision,sigla_voucher,observacion,estado) values (@empresa,@ano,@ncom,@ruc,@documento,@cuenta,@provision,@sigla_voucher,@observacion,@estado)", conx)
                                cmd188.Parameters.AddWithValue("@empresa", Trim(Label1.Text))
                                cmd188.Parameters.AddWithValue("@ano", Trim(TextBox1.Text))
                                cmd188.Parameters.AddWithValue("@ncom", Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value))
                                cmd188.Parameters.AddWithValue("@ruc", Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value))
                                cmd188.Parameters.AddWithValue("@documento", Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value))
                                cmd188.Parameters.AddWithValue("@cuenta", Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value))
                                cmd188.Parameters.AddWithValue("@provision", (DataGridView1.Rows(e.RowIndex).Cells(10).Value))
                                cmd188.Parameters.AddWithValue("@sigla_voucher", Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value))
                                cmd188.Parameters.AddWithValue("@observacion", "")
                                cmd188.Parameters.AddWithValue("@estado", 2)
                                cmd188.ExecuteNonQuery()
                                DataGridView1.Rows(e.RowIndex).Cells("tefe_3a").Value = 1
                                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.AliceBlue
                                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Red
                                Dim semana, smpe, semar, semar1 As Integer
                                Dim ano, ano1 As Integer
                                semana = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)

                                smpe = 53 - semana
                                If Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value) = "" Then
                                    semar = 1
                                    ano = "0"
                                Else
                                    semar = DatePart("ww", DataGridView1.Rows(e.RowIndex).Cells(6).Value, vbMonday, vbFirstFourDays)
                                    ano = Year(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
                                End If

                                If Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value) = "" Then
                                    semar1 = 1
                                    ano1 = "0"
                                Else
                                    semar1 = DatePart("ww", DataGridView1.Rows(e.RowIndex).Cells(7).Value, vbMonday, vbFirstFourDays)
                                    ano1 = Year(DataGridView1.Rows(e.RowIndex).Cells(7).Value)
                                End If
                                If Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value) = "" Then
                                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(15).Value) = "1" Then
                                        If semar < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano < Year(DateTimePicker1.Value) Then
                                            DataGridView1.Rows(e.RowIndex).Cells(11).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                        Else
                                            DataGridView1.Rows(e.RowIndex).Cells(11).Value = 0.0
                                            For i2 = 17 To smpe + 15

                                                If DataGridView1.Columns(i2).HeaderText = "SEM " & semar Then
                                                    DataGridView1.Rows(e.RowIndex).Cells(i2).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    Else

                                    End If

                                Else
                                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(15).Value) = "1" Then
                                        If semar1 < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano1 < Year(DateTimePicker1.Value) Then
                                            DataGridView1.Rows(e.RowIndex).Cells(11).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                        Else
                                            DataGridView1.Rows(e.RowIndex).Cells(11).Value = 0.0
                                            For i2 = 17 To smpe + 15

                                                If DataGridView1.Columns(i2).HeaderText = "SEM " & semar1 Then
                                                    DataGridView1.Rows(e.RowIndex).Cells(i2).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    Else

                                    End If

                                End If


                            End If
                        End If
                    End If
                    End If
            End If
        End If


    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Dim op As New Exportar
        op.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Dim semana, smpe As Integer
        semana = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        smpe = 53 - semana
        Dim pl As Integer
        pl = DataGridView2.Rows.Count
        Dim pl1 As Integer
        Dim can, can1 As Double
        pl1 = DataGridView1.Rows.Count

        For p1 = 1 To smpe

            For p2 = 13 To smpe + 11

                If Trim(DataGridView2.Columns(p1).HeaderText) = Trim(DataGridView1.Columns(p2).HeaderText) Then

                    For p3 = 0 To pl1 - 1
                        If Trim(DataGridView1.Rows(p3).Cells(9).Value) = "S/." Then
                            can = can + Val(DataGridView1.Rows(p3).Cells(p2).Value)
                        Else
                            can1 = can1 + Val(DataGridView1.Rows(p3).Cells(p2).Value)
                        End If
                    Next
                    DataGridView2.Rows(0).Cells(p1).Value = can
                    DataGridView2.Rows(1).Cells(p1).Value = can1
                Else
                    can = 0
                    can1 = 0
                End If
            Next

        Next
    End Sub
    Dim Rsr199 As SqlDataReader
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i1 As Integer
        Dim cco As String
        i1 = DataGridView1.Rows.Count
        For i = 0 To i1 - 1
            If DataGridView1.Rows(i).Cells(0).Value = True Then
                Dim cmd16 As New SqlCommand("update custom_vianny.dbo.cac3p00 set MONE=@MONE , f_program = @fecha, cancelar =@cancelar, fregistro =@fregistro where ncom_3a =@ncom_3 and ccia_3a =@ccia_3 and cperiodo_3a =@cperiodo_3 and fich_3a =@ruc AND ccom_3a =@ccom_3 and ndoc_3a =@ndoc_3a and cuen_3a =@cuenta", conx)
                cmd16.Parameters.AddWithValue("@fecha", Trim(DataGridView1.Rows(i).Cells(7).Value))
                cmd16.Parameters.AddWithValue("@ncom_3", Trim(DataGridView1.Rows(i).Cells(13).Value))
                cmd16.Parameters.AddWithValue("@ccia_3", Trim(Label1.Text))
                cmd16.Parameters.AddWithValue("@cperiodo_3", Trim(TextBox1.Text))
                cmd16.Parameters.AddWithValue("@ruc", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd16.Parameters.AddWithValue("@ndoc_3a", Trim(DataGridView1.Rows(i).Cells(4).Value))
                cmd16.Parameters.AddWithValue("@cancelar", Trim(DataGridView1.Rows(i).Cells(10).Value))
                cmd16.Parameters.AddWithValue("@cuenta", Trim(DataGridView1.Rows(i).Cells(1).Value))
                If Trim(DataGridView1.Rows(i).Cells(9).Value) = "S/." Then
                    cmd16.Parameters.AddWithValue("@MONE", 1)
                Else
                    cmd16.Parameters.AddWithValue("@MONE", 2)
                End If

                cmd16.Parameters.AddWithValue("@fregistro", Format(DateTimePicker1.Value.ToShortDateString, "Short Date"))
                Dim Rsr199 As SqlDataReader
                Dim sql1011 As String = "select ncom_3i from custom_vianny.dbo.cag3i00 where ccom_3i ='" + Trim(DataGridView1.Rows(i).Cells(12).Value) + "' and ccia_3i='" + Label1.Text + "'"
                Dim cmd1011 As New SqlCommand(sql1011, conx)
                Rsr199 = cmd1011.ExecuteReader()
                Rsr199.Read()
                cco = Rsr199(0)
                Rsr199.Close()
                cmd16.Parameters.AddWithValue("@ccom_3", Trim(cco))

                cmd16.ExecuteNonQuery()


            End If

        Next
        MsgBox("SE ACTUALIZO LA INFORMACION CON EXITO")
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged

        'DataGridView1.BeginEdit(True)

        'Dim J As Integer

        'J = DataGridView1.Rows.Count
        'For I = 0 To J - 1
        '    If Me.DataGridView1.CurrentRow.Index.ToString() = I Then

        '        DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.DarkCyan
        '        DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White
        '        'If Trim(DataGridView1.Rows(I).Cells(7).Value).Length > 0 Then
        '        '    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Red
        '        '    DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White
        '        'End If
        '    Else
        '        DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.Black
        '        DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.White
        '        'If Trim(DataGridView1.Rows(I).Cells(7).Value).Length > 0 Then
        '        '    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Red
        '        '    DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White
        '        'End If
        '    End If
        'Next
    End Sub

    Private WithEvents CellTextBox As DataGridViewTextBoxEditingControl

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        'buscar_cuenta()
        bs.Filter = String.Format("DOCUM LIKE '%{0}%'", TextBox6.Text)

        Dim semanaa, smpea, smpe, semana As Integer
        semanaa = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        semana = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        smpea = 53 - semanaa
        smpe = 53 - semana
        Dim pl As Integer
        pl = DataGridView2.Rows.Count
        Dim pl1, su, semar, semar1 As Integer
        Dim can, can1 As Double
        Dim ano, ano1 As Integer
        Dim va, va1, va2, va3 As Double


        su = DataGridView1.Rows.Count
        For i1 = 0 To su - 1
            If Trim(DataGridView1.Rows(i1).Cells(6).Value) = "" Then
                semar = 1
                ano = "0"
            Else
                semar = DatePart("ww", DataGridView1.Rows(i1).Cells(6).Value, vbMonday, vbFirstFourDays)
                ano = Year(DataGridView1.Rows(i1).Cells(6).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(7).Value) = "" Then
                semar1 = 1
                ano1 = "0"
            Else
                semar1 = DatePart("ww", DataGridView1.Rows(i1).Cells(7).Value, vbMonday, vbFirstFourDays)
                ano1 = Year(DataGridView1.Rows(i1).Cells(7).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(7).Value) = "" Then
                If semar < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano < Year(DateTimePicker1.Value) Then
                    DataGridView1.Rows(i1).Cells(11).Value = DataGridView1.Rows(i1).Cells(10).Value
                Else
                    DataGridView1.Rows(i1).Cells(11).Value = 0.0
                    For i2 = 17 To smpe + 15

                        If DataGridView1.Columns(i2).HeaderText = "SEM " & semar Then
                            DataGridView1.Rows(i1).Cells(i2).Value = DataGridView1.Rows(i1).Cells(10).Value
                            Exit For
                        End If
                    Next
                End If
            Else

                If semar1 < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano1 < Year(DateTimePicker1.Value) Then
                    DataGridView1.Rows(i1).Cells(11).Value = DataGridView1.Rows(i1).Cells(10).Value
                Else
                    DataGridView1.Rows(i1).Cells(11).Value = 0.0
                    For i2 = 17 To smpe + 15

                        If DataGridView1.Columns(i2).HeaderText = "SEM " & semar1 Then
                            DataGridView1.Rows(i1).Cells(i2).Value = DataGridView1.Rows(i1).Cells(10).Value
                            Exit For
                        End If
                    Next
                End If
            End If



            If Trim(DataGridView1.Rows(i1).Cells(9).Value) = "S/." Then
                va = va + Val(DataGridView1.Rows(i1).Cells(10).Value)
            Else
                va2 = va2 + Val(DataGridView1.Rows(i1).Cells(10).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(9).Value) = "S/." Then
                va1 = va1 + Val(DataGridView1.Rows(i1).Cells(11).Value)
            Else
                va3 = va3 + Val(DataGridView1.Rows(i1).Cells(11).Value)
            End If

            'If Trim(DataGridView1.Rows(i1).Cells(7).Value).Length > 0 And Trim(DataGridView1.Rows(i1).Cells(14).Value) = "" Then
            '    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
            '    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
            'End If

            If Trim(DataGridView1.Rows(i1).Cells(7).Value).Length > 0 And Trim(DataGridView1.Rows(i1).Cells(14).Value) = "" Then
                If Trim(DataGridView1.Rows(i1).Cells(7).Value) < DateTimePicker1.Value Then
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                Else
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.DarkCyan
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                End If
            Else
                If Trim(DataGridView1.Rows(i1).Cells(15).Value) = "" Then
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.DarkGray
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                Else

                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.White
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.Black
                End If
            End If
            If Trim(DataGridView1.Rows(i1).Cells(14).Value) = "1" Then
                DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Yellow
                DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.Black
                DataGridView1.Rows(i1).Cells(0).Value = True
            End If
            'If Trim(DataGridView1.Rows(i1).Cells(17).Value) = "1" Then
            '    DataGridView1.Rows(i1).Visible = False
            'End If
        Next
        For p1 = 1 To smpea

            For p2 = 17 To smpea + 15

                If Trim(DataGridView2.Columns(p1).HeaderText) = Trim(DataGridView1.Columns(p2).HeaderText) Then
                    pl1 = DataGridView1.Rows.Count
                    For p3 = 0 To pl1 - 1
                        If Trim(DataGridView1.Rows(p3).Cells(9).Value) = "S/." Then
                            can = can + Val(DataGridView1.Rows(p3).Cells(p2).Value)
                        Else
                            can1 = can1 + Val(DataGridView1.Rows(p3).Cells(p2).Value)
                        End If
                    Next
                    DataGridView2.Rows(0).Cells(p1).Value = can
                    DataGridView2.Rows(1).Cells(p1).Value = can1
                Else
                    can = 0
                    can1 = 0
                End If
            Next

        Next
        TextBox3.Text = va.ToString("N2")
        TextBox4.Text = va1.ToString("N2")
        TextBox7.Text = va2.ToString("N2")
        TextBox8.Text = va3.ToString("N2")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox5.Text = ""
        TextBox6.Text = ""
        Dim semana, smpe, semana1, smpe1 As Integer
        semana = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        semana1 = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        smpe = 53 - semana
        smpe1 = 53 - semana1
        DataGridView1.DataSource = ""

        For X = 1 To smpe
            DataGridView1.Columns.Remove(DataGridView1.Columns("SEM " & semana).HeaderText)
            semana = semana + 1
        Next

        DataGridView2.Rows.Clear()

        For X1 = 1 To smpe1
            DataGridView2.Columns.Remove(DataGridView2.Columns("SEM " & semana1).HeaderText)
            semana1 = semana1 + 1
        Next
        da.Clear()
        Button1.Enabled = True
        Button4.Enabled = False
        TextBox3.Text = 0.0
        TextBox4.Text = 0.0
        TextBox7.Text = 0.0
        TextBox8.Text = 0.0
        TextBox9.Text = 0.0
    End Sub


    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        CellTextBox = TryCast(e.Control, DataGridViewTextBoxEditingControl)
    End Sub

    Private Sub TextBox5_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        bs.Filter = String.Format("PROVEEDOR LIKE '%{0}%'", TextBox5.Text)

        Dim semanaa, smpea, smpe, semana As Integer
        semanaa = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        semana = DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays)
        smpea = 53 - semanaa
        smpe = 53 - semana
        Dim pl As Integer
        pl = DataGridView2.Rows.Count
        Dim pl1, su, semar, semar1 As Integer
        Dim can, can1 As Double
        Dim ano, ano1 As Integer
        Dim va, va1, va2, va3 As Double


        su = DataGridView1.Rows.Count
        For i1 = 0 To su - 1
            If Trim(DataGridView1.Rows(i1).Cells(6).Value) = "" Then
                semar = 1
                ano = "0"
            Else
                semar = DatePart("ww", DataGridView1.Rows(i1).Cells(6).Value, vbMonday, vbFirstFourDays)
                ano = Year(DataGridView1.Rows(i1).Cells(6).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(7).Value) = "" Then
                semar1 = 1
                ano1 = "0"
            Else
                semar1 = DatePart("ww", DataGridView1.Rows(i1).Cells(7).Value, vbMonday, vbFirstFourDays)
                ano1 = Year(DataGridView1.Rows(i1).Cells(7).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(7).Value) = "" Then
                If semar < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano < Year(DateTimePicker1.Value) Then
                    DataGridView1.Rows(i1).Cells(11).Value = DataGridView1.Rows(i1).Cells(10).Value
                Else
                    DataGridView1.Rows(i1).Cells(11).Value = 0.0
                    For i2 = 17 To smpe + 15

                        If DataGridView1.Columns(i2).HeaderText = "SEM " & semar Then
                            DataGridView1.Rows(i1).Cells(i2).Value = DataGridView1.Rows(i1).Cells(10).Value
                            Exit For
                        End If
                    Next
                End If
            Else

                If semar1 < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano1 < Year(DateTimePicker1.Value) Then
                    DataGridView1.Rows(i1).Cells(11).Value = DataGridView1.Rows(i1).Cells(10).Value
                Else
                    DataGridView1.Rows(i1).Cells(11).Value = 0.0
                    For i2 = 17 To smpe + 15

                        If DataGridView1.Columns(i2).HeaderText = "SEM " & semar1 Then
                            DataGridView1.Rows(i1).Cells(i2).Value = DataGridView1.Rows(i1).Cells(10).Value
                            Exit For
                        End If
                    Next
                End If
            End If



            If Trim(DataGridView1.Rows(i1).Cells(9).Value) = "S/." Then
                va = va + Val(DataGridView1.Rows(i1).Cells(10).Value)
            Else
                va2 = va2 + Val(DataGridView1.Rows(i1).Cells(10).Value)
            End If
            If Trim(DataGridView1.Rows(i1).Cells(9).Value) = "S/." Then
                va1 = va1 + Val(DataGridView1.Rows(i1).Cells(11).Value)
            Else
                va3 = va3 + Val(DataGridView1.Rows(i1).Cells(11).Value)
            End If

            'If Trim(DataGridView1.Rows(i1).Cells(7).Value).Length > 0 And Trim(DataGridView1.Rows(i1).Cells(14).Value) = "" Then
            '    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
            '    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
            'End If

            If Trim(DataGridView1.Rows(i1).Cells(7).Value).Length > 0 And Trim(DataGridView1.Rows(i1).Cells(14).Value) = "" Then
                If Trim(DataGridView1.Rows(i1).Cells(7).Value) < DateTimePicker1.Value Then
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                Else
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.DarkCyan
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                End If
            Else
                If Trim(DataGridView1.Rows(i1).Cells(15).Value) = "" Then
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.DarkGray
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                Else

                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.White
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.Black
                End If
            End If
            If Trim(DataGridView1.Rows(i1).Cells(14).Value) = "1" Then
                DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Yellow
                DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.Black
                DataGridView1.Rows(i1).Cells(0).Value = True
            End If
            'If Trim(DataGridView1.Rows(i1).Cells(17).Value) = "1" Then
            '    DataGridView1.Rows(i1).Visible = False
            'End If
        Next
        For p1 = 1 To smpea

            For p2 = 17 To smpea + 15

                If Trim(DataGridView2.Columns(p1).HeaderText) = Trim(DataGridView1.Columns(p2).HeaderText) Then
                    pl1 = DataGridView1.Rows.Count
                    For p3 = 0 To pl1 - 1
                        If Trim(DataGridView1.Rows(p3).Cells(9).Value) = "S/." Then
                            can = can + Val(DataGridView1.Rows(p3).Cells(p2).Value)
                        Else
                            can1 = can1 + Val(DataGridView1.Rows(p3).Cells(p2).Value)
                        End If
                    Next
                    DataGridView2.Rows(0).Cells(p1).Value = can
                    DataGridView2.Rows(1).Cells(p1).Value = can1
                Else
                    can = 0
                    can1 = 0
                End If
            Next

        Next
        TextBox3.Text = va.ToString("N2")
        TextBox4.Text = va1.ToString("N2")
        TextBox7.Text = va2.ToString("N2")
        TextBox8.Text = va3.ToString("N2")
    End Sub

    'Private Sub Button5_Click(sender As Object, e As EventArgs)
    '    Dim op As New Exportar2
    '    op.llenarExcel(DataGridView1)
    'End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown

    End Sub
End Class