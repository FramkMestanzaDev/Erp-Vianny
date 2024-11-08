Imports System.IO
Imports System.IO.StreamWriter
Public Class PDT_LIBROS
    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim folder As New FolderBrowserDialog
        Dim ruta As String = String.Empty

        If folder.ShowDialog = Windows.Forms.DialogResult.OK Then
            ruta = folder.SelectedPath
            TextBox2.Text = ruta
            TextBox4.Text = 2
        End If



    End Sub

    Private Sub PDT_LIBROS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim h As String
        ComboBox1.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        TextBox4.Text = 1
        ComboBox2.SelectedIndex = Month(DateTimePicker1.Value) - 1
        Select Case ComboBox2.Text
            Case "ENERO" : h = "01"
            Case "FEBRERO" : h = "02"
            Case "MARZO" : h = "03"
            Case "ABRIL" : h = "04"
            Case "MAYO" : h = "05"
            Case "JUNIO" : h = "06"
            Case "JULIO" : h = "07"
            Case "AGOSTO" : h = "08"
            Case "SETIEMBRE" : h = "09"
            Case "OCTUBRE" : h = "10"
            Case "NOVIEMBRE" : h = "11"
            Case "DICIEMBRE" : h = "12"
        End Select

        'TextBox2.Text = "Z:\CONTABILIDAD\USUARIO YOEL PAZ\PLE-PRUEBA" + "/" + "LE" + TextBox3.Text + TextBox1.Text + h + "00" + "140100001111" + ".txt"
    End Sub

    Function LibroDiario()
        Dim AR40, ar50, ar41, formato, correaltivo As String
        Dim ju As New vcontablidad
        Dim ju1 As New fcontailidad
        Dim linea3 As String = Nothing
        ar41 = 1
        Dim archivo5 As String
        Dim archivo6 As StreamWriter
        Dim h1 As String
        h1 = ""
        Select Case ComboBox2.Text
            Case "ENERO" : h1 = "01"
            Case "FEBRERO" : h1 = "02"
            Case "MARZO" : h1 = "03"
            Case "ABRIL" : h1 = "04"
            Case "MAYO" : h1 = "05"
            Case "JUNIO" : h1 = "06"
            Case "JULIO" : h1 = "07"
            Case "AGOSTO" : h1 = "08"
            Case "SETIEMBRE" : h1 = "09"
            Case "OCTUBRE" : h1 = "10"
            Case "NOVIEMBRE" : h1 = "11"
            Case "DICIEMBRE" : h1 = "12"
        End Select
        If Trim(TextBox4.Text) = "2" Then
            archivo5 = TextBox2.Text + "/" + "LE" + Trim(TextBox3.Text) + Trim(TextBox1.Text) + h1 + "00" + "050100001111" + ".txt"
            archivo6 = New StreamWriter(Trim(TextBox2.Text) + "/" + "LE" + Trim(TextBox3.Text) + Trim(TextBox1.Text) + h1 + "00" + "050100001111" + ".txt")
        Else
            archivo5 = Trim(TextBox2.Text)
            archivo6 = New StreamWriter(TextBox2.Text)
        End If
        For i = 0 To DataGridView3.Rows.Count - 1

            'If i >= 1 Then
            '    If Trim(DataGridView3.Item(1, i).Value) = Trim(DataGridView3.Item(1, i - 1).Value) Then
            '        AR40 = ar41 + 1
            '    Else
            '        AR40 = 1
            '    End If
            '    ar41 = AR40

            'End If
            'ar50 = "000" + ar41
            'If Trim(DataGridView3.Item(6, i).Value) = "" Then
            '    formato = ""
            'Else
            '    formato = Trim(DataGridView3.Item(6, i).Value) & Trim(DataGridView3.Item(2, i).Value) & Microsoft.VisualBasic.Right(ar50, 4)
            'End If

            'If Trim(DataGridView3.Item(4, i).Value) = "" Then
            '    correaltivo = "0"
            'Else
            '    correaltivo = Trim(DataGridView3.Item(4, i).Value)
            'End If

            linea3 = linea3 & Trim(DataGridView3.Item(0, i).Value) & "|" + vbCrLf


        Next

        archivo6.Write(linea3)
        archivo6.Close()
    End Function
    Function LibroMayor()
        Dim AR40, ar50, ar41, formato, correaltivo As String
        Dim ju As New vcontablidad
        Dim ju1 As New fcontailidad
        Dim linea3 As String = Nothing
        ar41 = 1
        Dim archivo5 As String
        Dim archivo6 As StreamWriter
        Dim h1 As String
        h1 = ""
        Select Case ComboBox2.Text
            Case "ENERO" : h1 = "01"
            Case "FEBRERO" : h1 = "02"
            Case "MARZO" : h1 = "03"
            Case "ABRIL" : h1 = "04"
            Case "MAYO" : h1 = "05"
            Case "JUNIO" : h1 = "06"
            Case "JULIO" : h1 = "07"
            Case "AGOSTO" : h1 = "08"
            Case "SETIEMBRE" : h1 = "09"
            Case "OCTUBRE" : h1 = "10"
            Case "NOVIEMBRE" : h1 = "11"
            Case "DICIEMBRE" : h1 = "12"
        End Select
        If TextBox4.Text = 2 Then
            archivo5 = TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h1 + "00" + "060100001111" + ".txt"
            archivo6 = New StreamWriter(TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h1 + "00" + "060100001111" + ".txt")
        Else
            archivo5 = TextBox2.Text
            archivo6 = New StreamWriter(TextBox2.Text)
        End If
        For i = 0 To DataGridView3.Rows.Count - 1
            linea3 = linea3 & Trim(DataGridView3.Item(0, i).Value) & "|" + vbCrLf
        Next

        archivo6.Write(linea3)
        archivo6.Close()
    End Function

    Dim dt As New DataTable
    Dim dt1, dt2, dt4 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FH As New flibros
        Dim JK As New vlibros
        Dim h As String
        h = ""
        If TextBox2.Text = "" Then
            MsgBox("NO EXISTE LA RUTA DONDE SE GUARDARA EL ARCHIVO")
        Else
            If ComboBox1.Text = "REGISTRO VENTAS" Then
                Select Case ComboBox3.Text
                    Case "VIANNY" : JK.gccia = "01"
                    Case "GRAUS" : JK.gccia = "02"
                    Case "DISEÑO" : JK.gccia = "03"
                End Select
                Select Case ComboBox2.Text
                    Case "ENERO" : h = "01"
                    Case "FEBRERO" : h = "02"
                    Case "MARZO" : h = "03"
                    Case "ABRIL" : h = "04"
                    Case "MAYO" : h = "05"
                    Case "JUNIO" : h = "06"
                    Case "JULIO" : h = "07"
                    Case "AGOSTO" : h = "08"
                    Case "SETIEMBRE" : h = "09"
                    Case "OCTUBRE" : h = "10"
                    Case "NOVIEMBRE" : h = "11"
                    Case "DICIEMBRE" : h = "12"
                End Select
                JK.gCPERIODO = TextBox1.Text

                JK.gfecha_ini = DateSerial(TextBox1.Text, h, 1)
                JK.gfecha_FIN = DateSerial(TextBox1.Text, h + 1, 0)

                dt = FH.Reporte_Ventas(JK)
                DataGridView1.DataSource = dt
                Dim archivo1 As String
                Dim archivo As StreamWriter
                Dim AR, ar1, ar2, ar3 As String
                Dim linea As String = Nothing
                Dim linea2 As String = Nothing
                Dim linea3 As String = Nothing
                Dim linea4 As String = Nothing
                Dim linea5 As String = Nothing
                Dim linea6 As String = Nothing
                If TextBox4.Text = 2 Then
                    archivo1 = TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h + "00" + "140100001111" + ".txt"
                    archivo = New StreamWriter(TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h + "00" + "140100001111" + ".txt")
                Else
                    archivo1 = TextBox2.Text
                    archivo = New StreamWriter(TextBox2.Text)
                End If



                With DataGridView1


                    For i = 0 To DataGridView1.Rows.Count - 1
                        If DataGridView1.Item(22, i).Value Is DBNull.Value Then
                            If Trim(DataGridView1.Item(17, i).Value) = "008" Then
                                If DataGridView1.Item(9, i).Value Is DBNull.Value Then
                                    AR = ""
                                Else
                                    AR = Trim(DataGridView1.Item(9, i).Value)
                                End If

                            Else
                                AR = "01/01/0001"
                            End If

                        Else
                            AR = Format(DataGridView1.Item(22, i).Value, "dd/MM/yyyy")
                        End If
                        If Convert.ToString(Trim(DataGridView1.Item(23, i).Value)) = "" Then
                            ar1 = "00"
                        Else
                            ar1 = Mid(Trim(DataGridView1.Item(23, i).Value), 2, 2)
                        End If
                        If Convert.ToString(Trim(DataGridView1.Item(24, i).Value)) = "" Then
                            ar2 = "-"
                        Else
                            ar2 = Trim(DataGridView1.Item(24, i).Value)
                        End If
                        If Convert.ToString(Trim(DataGridView1.Item(25, i).Value)) = "" Then
                            ar3 = "0"
                        Else
                            ar3 = Trim(DataGridView1.Item(25, i).Value).TrimStart("0")
                        End If
                        linea2 = linea2 & TextBox1.Text & h & "00" & "|" & Trim(DataGridView1.Item(27, i).Value) & Trim(DataGridView1.Item(6, i).Value) & "|" & "M" & Trim(DataGridView1.Item(27, i).Value) & Mid(Trim(DataGridView1.Item(6, i).Value), 1, 7) & "|" & Format(DataGridView1.Item(8, i).Value, "dd/MM/yyyy") & "|" & "" & "|" &
                        Mid(Trim(DataGridView1.Item(17, i).Value), 2, 2) & "|" & Trim(DataGridView1.Item(19, i).Value) & "|" & Trim(DataGridView1.Item(20, i).Value).TrimStart("0") & "|" & "" & "|" & Trim(DataGridView1.Item(12, i).Value).TrimStart("0") & "|" &
                        Trim(DataGridView1.Item(14, i).Value) & "|" & Trim(DataGridView1.Item(15, i).Value) & "|" & Trim(DataGridView1.Item(28, i).Value) & "|" & Trim(DataGridView1.Item(29, i).Value) & "|" & "0.00" & "|" &
                        Trim(DataGridView1.Item(31, i).Value) & "|" & "0.00" & "|" & "0.00" & "|" & "0.00" & "|" & "0.00" & "|" & "0.00" & "|" & "0.00" & "|" & "0.00" & "|" & "0.00" & "|" & Trim(DataGridView1.Item(32, i).Value) & "|" &
                         Trim(DataGridView1.Item(11, i).Value) & "|" & Mid(DataGridView1.Item(21, i).Value, 1, 5) & "|" & AR & "|" & ar1 & "|" & ar2 & "|" & ar3 & "|" & "" & "|" & "" & "|" & "" & "|" & Trim(DataGridView1.Item(26, i).Value) & "|" + vbCrLf


                    Next



                    archivo.Write(linea2)
                    archivo.Close()

                End With

                MsgBox("SE CREO EL ARCHIVO VENTAS CORRECTAMENTE")
                DataGridView1.DataSource = ""
            Else
                If ComboBox1.Text = "REGISTRO COMPRAS" Then
                    Select Case ComboBox3.Text
                        Case "VIANNY" : JK.gccia = "01"
                        Case "GRAUS" : JK.gccia = "02"
                        Case "DISEÑO" : JK.gccia = "03"
                    End Select
                    Select Case ComboBox2.Text
                        Case "ENERO" : h = "01"
                        Case "FEBRERO" : h = "02"
                        Case "MARZO" : h = "03"
                        Case "ABRIL" : h = "04"
                        Case "MAYO" : h = "05"
                        Case "JUNIO" : h = "06"
                        Case "JULIO" : h = "07"
                        Case "AGOSTO" : h = "08"
                        Case "SETIEMBRE" : h = "09"
                        Case "OCTUBRE" : h = "10"
                        Case "NOVIEMBRE" : h = "11"
                        Case "DICIEMBRE" : h = "12"
                    End Select
                    JK.gCPERIODO = TextBox1.Text

                    JK.gfecha_ini = DateSerial(TextBox1.Text, h, 1)
                    JK.gfecha_FIN = DateSerial(TextBox1.Text, h + 1, 0)

                    dt1 = FH.Reporte_Compras(JK)
                    DataGridView2.DataSource = dt1
                    Dim archivo3 As String
                    Dim archivo2 As StreamWriter
                    Dim AR, ar1, ar2, ar3, AR40, ar5, numdet, ar50, numdet1, ar41, AR24 As String
                    Dim ju As New vcontablidad
                    Dim ju1 As New fcontailidad
                    Dim linea3 As String = Nothing
                    ar41 = 1
                    If TextBox4.Text = 2 Then
                        archivo3 = TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h + "00" + "080100001111" + ".txt"
                        archivo2 = New StreamWriter(TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h + "00" + "080100001111" + ".txt")
                    Else
                        archivo3 = TextBox2.Text
                        archivo2 = New StreamWriter(TextBox2.Text)
                    End If



                    With DataGridView2


                        For i = 0 To DataGridView2.Rows.Count - 1
                            If DataGridView2.Item(13, i).Value = "014" Then
                                AR24 = Format(DataGridView2.Item(5, i).Value, "dd/MM/yyyy")
                            Else
                                AR24 = ""
                            End If

                            If DataGridView2.Item(20, i).Value Is DBNull.Value Then
                                AR = "01/01/0001"
                            Else
                                AR = Format(DataGridView2.Item(20, i).Value, "dd/MM/yyyy")
                            End If
                            If Convert.ToString(Trim(DataGridView2.Item(21, i).Value)) = "" Then
                                ar1 = "00"
                            Else
                                ar1 = Mid(Trim(DataGridView2.Item(21, i).Value), 2, 2)
                            End If
                            If Convert.ToString(Trim(DataGridView2.Item(22, i).Value)) = "" Then
                                ar2 = "-"
                            Else
                                ar2 = Trim(DataGridView2.Item(22, i).Value)
                            End If
                            If Convert.ToString(Trim(DataGridView2.Item(23, i).Value)) = "" Then
                                ar3 = "-"
                            Else
                                ar3 = Trim(DataGridView2.Item(23, i).Value).TrimStart("0")
                            End If

                            If DataGridView2.Item(24, i).Value Is DBNull.Value Then
                                ar5 = "01/01/0001"
                            Else
                                ar5 = Format(DataGridView2.Item(24, i).Value, "dd/MM/yyyy")
                            End If
                            ju.gfactura = Trim(DataGridView2.Item(15, i).Value) & Trim(DataGridView2.Item(16, i).Value)
                            ju.gdoc = Trim(DataGridView2.Item(13, i).Value)
                            ju.gncom = Trim(DataGridView2.Item(2, i).Value)
                            ju.gruc = Trim(DataGridView2.Item(10, i).Value)
                            Select Case Trim(ComboBox3.Text)
                                Case "VIANNY" : ju.gcia = "01"
                                Case "GRAUS" : ju.gcia = "02"
                                Case "DISEÑO" : ju.gcia = "03"
                            End Select

                            numdet = ju1.buscar_numdet(ju)
                            If numdet = 0 Then
                                numdet1 = 0
                            Else
                                numdet1 = numdet
                            End If
                            If Month(DataGridView2.Item(3, i).Value) = h Then
                                ar50 = "1"
                            Else
                                ar50 = "6"
                            End If
                            If i >= 1 Then
                                If Trim(DataGridView2.Item(1, i).Value) = Trim(DataGridView2.Item(1, i - 1).Value) Then


                                    AR40 = ar41 + 1

                                Else
                                    AR40 = 1

                                End If
                                ar41 = AR40

                            End If
                            Dim KL, kl1 As String
                            If Mid(Trim(DataGridView2.Item(13, i).Value), 2, 2) = "50" Or Mid(Trim(DataGridView2.Item(13, i).Value), 2, 2) = "52" Then
                                KL = "2021"
                                kl1 = Trim(DataGridView2.Item(15, i).Value).TrimStart("0")
                            Else
                                KL = 0
                                kl1 = Trim(DataGridView2.Item(15, i).Value)
                            End If

                            linea3 = linea3 & TextBox1.Text & h & "00" & "|" & Trim(DataGridView2.Item(1, i).Value) & "|" & "M" & Trim(DataGridView2.Item(1, i).Value) & ar41 & "|" &
                        Format(DataGridView2.Item(3, i).Value, "dd/MM/yyyy") & "|" & AR24 & "|" & Mid(Trim(DataGridView2.Item(13, i).Value), 2, 2) & "|" & kl1 & "|" & KL & "|" &
                        Trim(DataGridView2.Item(16, i).Value).TrimStart("0") & "|" & "" & "|" & Trim(DataGridView2.Item(8, i).Value).TrimStart("0") & "|" & Trim(DataGridView2.Item(10, i).Value) & "|" &
                            Trim(DataGridView2.Item(11, i).Value) & "|" & Trim(DataGridView2.Item(43, i).Value) & "|" & Trim(DataGridView2.Item(48, i).Value) & "|" & "0.00" & "|" & "0.00" & "|" & "0.00" & "|" & "0.00" & "|" & Trim(DataGridView2.Item(46, i).Value) & "|" & "0.00" & "|" & "0.00" & "|" & "0.00" & "|" &
                            Trim(DataGridView2.Item(53, i).Value) & "|" & Trim(DataGridView2.Item(7, i).Value) & "|" & Mid(DataGridView2.Item(17, i).Value, 1, 5) & "|" & ar5 & "|" & ar1 & "|" & ar2 & "|" & "" & "|" & ar3 & "|" & AR & "|" &
                               numdet1 & "|" & "" & "|" & Trim(DataGridView2.Item(26, i).Value) & "|" & "" & "|" & "1" & "|" & "" & "|" & "" & "|" & "" & "|" & "1" & "|" & ar50 & "|" + vbCrLf


                        Next



                        archivo2.Write(linea3)
                        archivo2.Close()

                    End With
                    MsgBox("SE CREO EL ARCHIVO COMPRAS CORRECTAMENTE")
                    DataGridView2.DataSource = ""
                Else
                    If ComboBox1.Text = "REGISTRO COMPRAS(NO DOMICILIADO)	" Then
                        Dim h1 As String
                        h1 = ""
                        Select Case ComboBox2.Text
                            Case "ENERO" : h1 = "01"
                            Case "FEBRERO" : h1 = "02"
                            Case "MARZO" : h1 = "03"
                            Case "ABRIL" : h1 = "04"
                            Case "MAYO" : h1 = "05"
                            Case "JUNIO" : h1 = "06"
                            Case "JULIO" : h1 = "07"
                            Case "AGOSTO" : h1 = "08"
                            Case "SETIEMBRE" : h1 = "09"
                            Case "OCTUBRE" : h1 = "10"
                            Case "NOVIEMBRE" : h1 = "11"
                            Case "DICIEMBRE" : h1 = "12"
                        End Select
                        Dim archivo1 As String
                        Dim archivo As StreamWriter
                        If TextBox4.Text = 2 Then
                            archivo1 = TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h1 + "00" + "080200001011" + ".txt"
                            archivo = New StreamWriter(TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h1 + "00" + "080200001011" + ".txt")
                        Else
                            archivo1 = TextBox2.Text
                            archivo = New StreamWriter(TextBox2.Text)

                        End If
                        archivo.Close()
                        MsgBox("SE CREO EL ARCHIVO VENTAS CORRECTAMENTE")
                    Else
                        If ComboBox1.Text = "LIBRO DIARIO" Then
                            Dim h1 As String
                            h1 = ""
                            Select Case ComboBox2.Text
                                Case "ENERO" : h1 = "01"
                                Case "FEBRERO" : h1 = "02"
                                Case "MARZO" : h1 = "03"
                                Case "ABRIL" : h1 = "04"
                                Case "MAYO" : h1 = "05"
                                Case "JUNIO" : h1 = "06"
                                Case "JULIO" : h1 = "07"
                                Case "AGOSTO" : h1 = "08"
                                Case "SETIEMBRE" : h1 = "09"
                                Case "OCTUBRE" : h1 = "10"
                                Case "NOVIEMBRE" : h1 = "11"
                                Case "DICIEMBRE" : h1 = "12"
                            End Select


                            JK.gCPERIODO = Trim(TextBox1.Text)
                            Select Case ComboBox3.Text
                                Case "VIANNY" : JK.gccia = "01"
                                Case "GRAUS" : JK.gccia = "02"
                                Case "DISEÑO" : JK.gccia = "03"
                            End Select
                            JK.gmes = h1

                            dt2 = FH.Libro_Diario(JK)
                            DataGridView3.DataSource = dt2

                            LibroDiario()

                            MsgBox("SE CREO EL ARCHIVO COMPRAS CORRECTAMENTE")
                            DataGridView3.DataSource = ""
                            dt2.Clear()
                        Else
                            If ComboBox1.Text = "LIBRO MAYOR" Then
                                Dim h1 As String
                                h1 = ""
                                Select Case ComboBox2.Text
                                    Case "ENERO" : h1 = "01"
                                    Case "FEBRERO" : h1 = "02"
                                    Case "MARZO" : h1 = "03"
                                    Case "ABRIL" : h1 = "04"
                                    Case "MAYO" : h1 = "05"
                                    Case "JUNIO" : h1 = "06"
                                    Case "JULIO" : h1 = "07"
                                    Case "AGOSTO" : h1 = "08"
                                    Case "SETIEMBRE" : h1 = "09"
                                    Case "OCTUBRE" : h1 = "10"
                                    Case "NOVIEMBRE" : h1 = "11"
                                    Case "DICIEMBRE" : h1 = "12"
                                End Select
                                Dim archivo5 As String
                                Dim archivo6 As StreamWriter
                                Dim AR40, ar50, ar41, formato, correaltivo As String
                                Dim ju As New vcontablidad
                                Dim ju1 As New fcontailidad
                                Dim linea3 As String = Nothing
                                ar41 = 1

                                JK.gCPERIODO = Trim(TextBox1.Text)
                                Select Case ComboBox3.Text
                                    Case "VIANNY" : JK.gccia = "01"
                                    Case "GRAUS" : JK.gccia = "02"
                                    Case "DISEÑO" : JK.gccia = "03"
                                End Select
                                JK.gmes = h1

                                dt2 = FH.Libro_Diario(JK)
                                DataGridView3.DataSource = dt2
                                LibroMayor()
                                'If TextBox4.Text = 2 Then
                                '    archivo5 = TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h1 + "00" + "060100001111" + ".txt"
                                '    archivo6 = New StreamWriter(TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h1 + "00" + "060100001111" + ".txt")
                                'Else
                                '    archivo5 = TextBox2.Text
                                '    archivo6 = New StreamWriter(TextBox2.Text)
                                'End If




                                'For i = 0 To DataGridView3.Rows.Count - 1

                                '    ' 'AR24 = Format(DataGridView3.Item(12, i).Value, "dd/MM/yyyy")
                                '    ' 'AR25 = Format(DataGridView3.Item(13, i).Value, "dd/MM/yyyy")
                                '    ' 'AR26 = Format(DataGridView3.Item(14, i).Value, "dd/MM/yyyy")

                                '    ' If i >= 1 Then
                                '    '     If Trim(DataGridView3.Item(1, i).Value) = Trim(DataGridView3.Item(1, i - 1).Value) Then


                                '    '         AR40 = ar41 + 1

                                '    '     Else
                                '    '         AR40 = 1

                                '    '     End If
                                '    '     ar41 = AR40

                                '    ' End If

                                '    ' ar50 = "000" + ar41
                                '    ' If Trim(DataGridView3.Item(19, i).Value) = "" Then
                                '    '     formato = ""
                                '    ' Else
                                '    '     formato = Trim(DataGridView3.Item(19, i).Value) & Trim(DataGridView3.Item(2, i).Value) & Microsoft.VisualBasic.Right(ar50, 4)
                                '    ' End If
                                '    ' If Trim(DataGridView3.Item(11, i).Value) = "" Then
                                '    '     correaltivo = "0"
                                '    ' Else
                                '    '     correaltivo = Trim(DataGridView3.Item(11, i).Value)
                                '    ' End If


                                '    ' linea3 = linea3 & Trim(DataGridView3.Item(0, i).Value) & "|" & Trim(DataGridView3.Item(1, i).Value) & "|" & Trim(DataGridView3.Item(2, i).Value) & Microsoft.VisualBasic.Right(ar50, 4) & "|" & Trim(DataGridView3.Item(3, i).Value) & "|" & Trim(DataGridView3.Item(4, i).Value) & "|" &
                                '    ' Trim(DataGridView3.Item(5, i).Value) & "|" & Trim(DataGridView3.Item(6, i).Value) & "|" & Trim(DataGridView3.Item(7, i).Value) & "|" & Trim(DataGridView3.Item(8, i).Value) & "|" & Trim(DataGridView3.Item(9, i).Value) & "|" & Trim(DataGridView3.Item(10, i).Value) & "|" &
                                '    'correaltivo & "|" & Format(DataGridView3.Item(12, i).Value, "dd/MM/yyyy") & "|" & Format(DataGridView3.Item(13, i).Value, "dd/MM/yyyy") & "|" & Format(DataGridView3.Item(14, i).Value, "dd/MM/yyyy") & "|" & Trim(DataGridView3.Item(15, i).Value) & "|" & Trim(DataGridView3.Item(16, i).Value) & "|" &
                                '    ' (DataGridView3.Item(17, i).Value) & "|" & (DataGridView3.Item(18, i).Value) & "|" & formato & "|" & Trim(DataGridView3.Item(20, i).Value) & "|" + vbCrLf


                                'Next



                                'archivo6.Write(linea3)
                                '    archivo6.Close()


                                MsgBox("SE CREO EL ARCHIVO COMPRAS CORRECTAMENTE")
                                DataGridView3.DataSource = ""
                                dt2.Clear()
                            Else
                                If ComboBox1.Text = "LIBRO DIARIO 5.3" Then
                                    Dim h1 As String
                                    h1 = ""
                                    Select Case ComboBox2.Text
                                        Case "ENERO" : h1 = "01"
                                        Case "FEBRERO" : h1 = "02"
                                        Case "MARZO" : h1 = "03"
                                        Case "ABRIL" : h1 = "04"
                                        Case "MAYO" : h1 = "05"
                                        Case "JUNIO" : h1 = "06"
                                        Case "JULIO" : h1 = "07"
                                        Case "AGOSTO" : h1 = "08"
                                        Case "SETIEMBRE" : h1 = "09"
                                        Case "OCTUBRE" : h1 = "10"
                                        Case "NOVIEMBRE" : h1 = "11"
                                        Case "DICIEMBRE" : h1 = "12"
                                    End Select
                                    Dim archivo7 As String
                                    Dim archivo8 As StreamWriter
                                    'Dim AR40, ar50, ar41, formato, correaltivo As String
                                    Dim ju As New vcontablidad
                                    Dim ju1 As New fcontailidad
                                    Dim linea3 As String = Nothing
                                    'ar41 = 1

                                    JK.gCPERIODO = Trim(TextBox1.Text)
                                    Select Case ComboBox3.Text
                                        Case "VIANNY" : JK.gccia = "01"
                                        Case "GRAUS" : JK.gccia = "02"
                                        Case "DISEÑO" : JK.gccia = "03"
                                    End Select
                                    JK.gmes = h1

                                    dt4 = FH.Libro_Diario5_3(JK)
                                    DataGridView4.DataSource = dt4

                                    If TextBox4.Text = 2 Then
                                        archivo7 = TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h1 + "00" + "050300001111" + ".txt"
                                        archivo8 = New StreamWriter(TextBox2.Text + "/" + "LE" + TextBox3.Text + TextBox1.Text + h1 + "00" + "050300001111" + ".txt")
                                    Else
                                        archivo7 = TextBox2.Text
                                        archivo8 = New StreamWriter(TextBox2.Text)
                                    End If

                                    With DataGridView4


                                        For i = 0 To DataGridView4.Rows.Count - 1


                                            linea3 = linea3 & Trim(DataGridView4.Item(0, i).Value) & "|" & Trim(DataGridView4.Item(1, i).Value) & "|" & Trim(DataGridView4.Item(2, i).Value) & "|" & Trim(DataGridView4.Item(3, i).Value) & "|" & Trim(DataGridView4.Item(4, i).Value) & "|" &
                                            Trim(DataGridView4.Item(5, i).Value) & "|" & Trim(DataGridView4.Item(6, i).Value) & "|" & Trim(DataGridView4.Item(7, i).Value) & "|" + vbCrLf


                                        Next



                                        archivo8.Write(linea3)
                                        archivo8.Close()

                                    End With
                                    MsgBox("SE CREO EL ARCHIVO COMPRAS CORRECTAMENTE")
                                    DataGridView4.DataSource = ""
                                    dt4.Clear()
                                Else
                                    If ComboBox1.Text = "LIBRO RETENCIONES" Then
                                        Dim h1 As String
                                        h1 = ""
                                        Select Case ComboBox2.Text
                                            Case "ENERO" : h1 = "01"
                                            Case "FEBRERO" : h1 = "02"
                                            Case "MARZO" : h1 = "03"
                                            Case "ABRIL" : h1 = "04"
                                            Case "MAYO" : h1 = "05"
                                            Case "JUNIO" : h1 = "06"
                                            Case "JULIO" : h1 = "07"
                                            Case "AGOSTO" : h1 = "08"
                                            Case "SETIEMBRE" : h1 = "09"
                                            Case "OCTUBRE" : h1 = "10"
                                            Case "NOVIEMBRE" : h1 = "11"
                                            Case "DICIEMBRE" : h1 = "12"
                                        End Select
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        If ComboBox3.Text = "VIANNY" Then
            TextBox3.Text = "20508740361"
        Else
            If ComboBox3.Text = "GRAUS" Then
                TextBox3.Text = "20459785834"
            Else
                If ComboBox3.Text = "DISEÑO" Then
                    TextBox3.Text = "20513067241"
                End If
            End If
        End If

        'TextBox2.Text = "Z:\CONTABILIDAD\USUARIO YOEL PAZ\PLE-PRUEBA" + "/" + "LE" + TextBox3.Text + TextBox1.Text + h + "00" + "080100001111" + ".txt"
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim h As String
        Select Case ComboBox2.Text
            Case "ENERO" : h = "01"
            Case "FEBRERO" : h = "02"
            Case "MARZO" : h = "03"
            Case "ABRIL" : h = "04"
            Case "MAYO" : h = "05"
            Case "JUNIO" : h = "06"
            Case "JULIO" : h = "07"
            Case "AGOSTO" : h = "08"
            Case "SETIEMBRE" : h = "09"
            Case "OCTUBRE" : h = "10"
            Case "NOVIEMBRE" : h = "11"
            Case "DICIEMBRE" : h = "12"
        End Select
        'MsgBox(ComboBox1.Text)
        'If ComboBox1.Text = "REGISTRO COMPRAS" Then
        '    TextBox2.Text = "Z:\CONTABILIDAD\USUARIO YOEL PAZ\PLE-PRUEBA" + "/" + "LE" + TextBox3.Text + TextBox1.Text + h + "00" + "080100001111" + ".txt"
        'Else
        '    If ComboBox1.Text = "REGISTRO VENTAS" Then
        '        TextBox2.Text = "Z:\CONTABILIDAD\USUARIO YOEL PAZ\PLE-PRUEBA" + "/" + "LE" + TextBox3.Text + TextBox1.Text + h + "00" + "140100001111" + ".txt"
        '    Else
        '        TextBox2.Text = "Z:\CONTABILIDAD\USUARIO YOEL PAZ\PLE-PRUEBA" + "/" + "LE" + TextBox3.Text + TextBox1.Text + h + "00" + "080200001111" + ".txt"
        '    End If


        'End If
    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click

    End Sub
End Class