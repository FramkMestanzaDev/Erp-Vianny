Imports System.Data.SqlClient

Public Class Form_cuentarxcobrar
    Public conx As SqlConnection
    Public conn As SqlDataAdapter

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
    Dim da As New DataTable
    Dim bs As New BindingSource()
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = ""
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("exec cuentasxcobrar '" + Trim(TextBox1.Text) + "'", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs

            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(14).Visible = False
            DataGridView1.Columns(8).Visible = False
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
        For i = 13 To smpe + 12
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
                If semar < DatePart("ww", DateTimePicker1.Value, vbMonday, vbFirstFourDays) Or ano < Year(DateTimePicker1.Value) Then
                    DataGridView1.Rows(i1).Cells(11).Value = DataGridView1.Rows(i1).Cells(10).Value
                Else
                    DataGridView1.Rows(i1).Cells(11).Value = 0.0
                    For i2 = 15 To smpe + 14

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
                    For i2 = 15 To smpe + 14

                        If DataGridView1.Columns(i2).HeaderText = "SEM " & semar1 Then
                            DataGridView1.Rows(i1).Cells(i2).Value = DataGridView1.Rows(i1).Cells(10).Value
                            Exit For
                        End If
                    Next
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

            For p2 = 15 To smpea + 14

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
        Dim go, go1 As Double
        Lf = DataGridView1.Rows.Count
        For k = 0 To Lf - 1

            If Trim(DataGridView1.Rows(k).Cells(7).Value).Length > 0 Then
                If Trim(DataGridView1.Rows(k).Cells(7).Value) < DateTimePicker1.Value Then
                    DataGridView1.Rows(k).DefaultCellStyle.BackColor = Color.DarkCyan
                    DataGridView1.Rows(k).DefaultCellStyle.ForeColor = Color.White
                Else
                    DataGridView1.Rows(k).DefaultCellStyle.BackColor = Color.Red
                    DataGridView1.Rows(k).DefaultCellStyle.ForeColor = Color.White
                End If

            End If
            If Trim(DataGridView1.Rows(k).Cells(14).Value) = "1" Then
                DataGridView1.Rows(k).DefaultCellStyle.BackColor = Color.Yellow
                DataGridView1.Rows(k).DefaultCellStyle.ForeColor = Color.Black
                DataGridView1.Rows(k).Cells(0).Value = True
            End If
            If Trim(DataGridView1.Rows(k).Cells(10).Value) = 0.00 Then
                DataGridView1.Rows(k).Visible = False
            End If


            If Trim(DataGridView1.Rows(k).Cells(7).Value).Length > 0 Then
                If Trim(DataGridView1.Rows(k).Cells(9).Value) = "S/." Then
                    go = go + Val(DataGridView1.Rows(k).Cells(10).Value)
                Else
                    go1 = go1 + Val(DataGridView1.Rows(k).Cells(10).Value)
                End If
            End If

        Next
        TextBox9.Text = go.ToString("N2")
        TextBox13.Text = go1.ToString("N2")
        Button1.Enabled = False
        Button4.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim op As New Exportar
        op.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i1 As Integer
        Dim cco As String
        i1 = DataGridView1.Rows.Count
        For i = 0 To i1 - 1
            If DataGridView1.Rows(i).Cells(0).Value = True Then
                Dim cmd16 As New SqlCommand("update custom_vianny.dbo.cac3p00 set MONE=@MONE ,f_program = @fecha,cancelar =@cancelar,fregistro=@fregistro where ncom_3a =@ncom_3 and ccia_3a =@ccia_3 and cperiodo_3a =@cperiodo_3 and fich_3a =@ruc AND ccom_3a =@ccom_3 and ndoc_3a =@ndoc_3a and cuen_3a =@cuenta", conx)
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

        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

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
                    For i2 = 15 To smpe + 14

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
                    For i2 = 15 To smpe + 14

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

            If Trim(DataGridView1.Rows(i1).Cells(7).Value).Length > 0 And Trim(DataGridView1.Rows(i1).Cells(14).Value) = "" Then
                DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
                DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
            End If
            If Trim(DataGridView1.Rows(i1).Cells(14).Value) = "1" Then
                DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Yellow
                DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.Black
                DataGridView1.Rows(i1).Cells(0).Value = True
            End If

        Next
        For p1 = 1 To smpea

            For p2 = 15 To smpea + 14

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
        Dim Lf As Integer
        Dim go, go1 As Double
        Lf = DataGridView1.Rows.Count
        For k = 0 To Lf - 1
            If Trim(DataGridView1.Rows(k).Cells(7).Value).Length > 0 Then
                If Trim(DataGridView1.Rows(k).Cells(9).Value) = "S/." Then
                    go = go + Val(DataGridView1.Rows(k).Cells(10).Value)
                Else
                    go1 = go1 + Val(DataGridView1.Rows(k).Cells(10).Value)
                End If
            End If

        Next
        TextBox9.Text = go.ToString("N2")
        TextBox13.Text = go1.ToString("N2")
        TextBox3.Text = va.ToString("N2")
        TextBox4.Text = va1.ToString("N2")
        TextBox7.Text = va2.ToString("N2")
        TextBox8.Text = va3.ToString("N2")
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "F_PROGRA" Then
            If DataGridView1.Rows(e.RowIndex).Cells(0).Value = True Then
                Dim num1, num2 As Integer
                num1 = e.RowIndex.ToString
                num2 = e.ColumnIndex.ToString

                F_PROGRAMACION.Label3.Text = num1
                F_PROGRAMACION.Label2.Text = 2
                F_PROGRAMACION.ShowDialog()
            End If

        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
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
                    For i2 = 15 To smpe + 14

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
                    For i2 = 15 To smpe + 14

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

            If Trim(DataGridView1.Rows(i1).Cells(7).Value).Length > 0 And Trim(DataGridView1.Rows(i1).Cells(14).Value) = "" Then
                DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
                DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
            End If
            If Trim(DataGridView1.Rows(i1).Cells(14).Value) = "1" Then
                DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Yellow
                DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.Black
                DataGridView1.Rows(i1).Cells(0).Value = True
            End If

        Next
        For p1 = 1 To smpea

            For p2 = 15 To smpea + 14

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
        Dim Lf As Integer
        Dim go, go1 As Double
        Lf = DataGridView1.Rows.Count
        For k = 0 To Lf - 1
            If Trim(DataGridView1.Rows(k).Cells(7).Value).Length > 0 Then
                If Trim(DataGridView1.Rows(k).Cells(9).Value) = "S/." Then
                    go = go + Val(DataGridView1.Rows(k).Cells(10).Value)
                Else
                    go1 = go1 + Val(DataGridView1.Rows(k).Cells(10).Value)
                End If
            End If

        Next
        TextBox9.Text = go.ToString("N2")
        TextBox13.Text = go1.ToString("N2")
        TextBox3.Text = va.ToString("N2")
        TextBox4.Text = va1.ToString("N2")
        TextBox7.Text = va2.ToString("N2")
        TextBox8.Text = va3.ToString("N2")
    End Sub
End Class