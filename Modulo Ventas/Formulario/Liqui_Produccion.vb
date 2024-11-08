Imports System.Data.SqlClient
Public Class Liqui_Produccion
    Public conx As SqlConnection
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Dim Rsr2, Rsr205, Rsr24, Rsr29, Rsr26 As SqlDataReader

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.F1 Then
            TRABAJADORES.Label2.Text = 6
            TRABAJADORES.ShowDialog()
        End If
    End Sub

    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.F1 Then
            TRABAJADORES.Label2.Text = 7
            TRABAJADORES.ShowDialog()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Trim(TextBox4.Text).Length
                Case "1" : TextBox4.Text = "01" & "0000000" & TextBox4.Text
                Case "2" : TextBox4.Text = "01" & "000000" & TextBox4.Text
                Case "3" : TextBox4.Text = "01" & "00000" & TextBox4.Text
                Case "4" : TextBox4.Text = "01" & "0000" & TextBox4.Text
                Case "5" : TextBox4.Text = "01" & "000" & TextBox4.Text
                Case "6" : TextBox4.Text = "01" & "00" & TextBox4.Text
                Case "7" : TextBox4.Text = "01" & "0" & TextBox4.Text
                Case "8" : TextBox4.Text = "01" & TextBox4.Text

            End Select

            abrir()
            Dim sql102 As String = "select ncom_3,linea_3,a.nomb_17, q.tallador_3,q.cants_3,q.program_3,a.dmarca_17,q.alterno_3  from custom_vianny.dbo.qag0300 q
            left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
            where ncom_3 ='" + Trim(TextBox4.Text) + "' and q.ccia ='" + Trim(Label15.Text) + "'
            "
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()


            If Rsr2.Read() = True Then
                DataGridView1.Rows.Add(2)
                TextBox5.Text = Rsr2(2)
                TextBox2.Text = Rsr2(5)
                TextBox6.Text = Rsr2(6)
                TextBox7.Text = Rsr2(4)
                TextBox8.Text = Rsr2(4)
                TextBox9.Text = Rsr2(7)
                DataGridView1.Rows(0).Cells(0).Value = "DESPACHO AL CLIENTE"
                DataGridView1.Rows(0).Cells(11).Value = Rsr2(3)
                DataGridView1.Rows(1).Cells(0).Value = "DESPACHO AL TIENDA"
                DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(1)
                DataGridView1.BeginEdit(True)
                Rsr2.Close()
            End If
            Rsr2.Close()

            Dim t1, t2, t3, t4, t5, t6, t7, t8, t9, t10 As Integer
            t1 = 0
            t2 = 0
            t3 = 0
            t4 = 0
            t5 = 0
            t6 = 0
            t7 = 0
            t8 = 0
            t9 = 0
            t10 = 0
            Dim sql10276 As String = "SELECT sum(cant1_3b) as T1,sum(cant2_3b) as T2,sum(cant3_3b) as T3,sum(cant4_3b) as T4,sum(cant5_3b) as T5,sum(cant6_3b) as T6,sum(cant7_3b) as T7,sum(cant8_3b) as T8,sum(cant9_3b)as T9,sum(cant10_3b)as T10 FROM custom_vianny.dbo.Mat030f  where pedido_3b ='0100022017' and flag_3b ='1' and ccia_3b ='01'"
            Dim cmd10276 As New SqlCommand(sql10276, conx)
            Rsr26 = cmd10276.ExecuteReader()

            If Rsr26.Read() = True Then
                t1 = Trim(Rsr26(0))
                t2 = Trim(Rsr26(1))
                t3 = Trim(Rsr26(2))
                t4 = Trim(Rsr26(3))
                t5 = Trim(Rsr26(4))
                t6 = Trim(Rsr26(5))
                t7 = Trim(Rsr26(6))
                t8 = Trim(Rsr26(7))
                t9 = Trim(Rsr26(8))
                t10 = Trim(Rsr26(9))
            End If
            Rsr26.Close()



            Dim sql10235 As String = "SELECT A.*,
                            		  ISNULL(B.Dele,'') AS NTalla1,
                            		  ISNULL(C.Dele,'') AS NTalla2,
                            		  ISNULL(D.Dele,'') AS NTalla3,
                            		  ISNULL(E.Dele,'') AS NTalla4,
                            		  ISNULL(F.Dele,'') AS NTalla5,
                            		  ISNULL(G.Dele,'') AS NTalla6,
                            		  ISNULL(H.Dele,'') AS NTalla7,
                            		  ISNULL(I.Dele,'') AS NTalla8,
                            		  ISNULL(J.Dele,'') AS NTalla9,
                            		  ISNULL(K.Dele,'') AS NTalla10
                             FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
                             				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
                             				 LEFT JOIN custom_vianny.dbo.Tab0100 C
                             				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
                             				 LEFT JOIN custom_vianny.dbo.Tab0100 D
                             				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
                             				 LEFT JOIN custom_vianny.dbo.Tab0100 E
                             				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
                             				 LEFT JOIN custom_vianny.dbo.Tab0100 F
                             				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
                             				 LEFT JOIN custom_vianny.dbo.Tab0100 G
                             				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
                             				 LEFT JOIN custom_vianny.dbo.Tab0100 H
                             				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
                             				 LEFT JOIN custom_vianny.dbo.Tab0100 I
                             				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
                             				 LEFT JOIN custom_vianny.dbo.Tab0100 J
                             				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
                             				 LEFT JOIN custom_vianny.dbo.Tab0100 K
                             				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
                             Where A.CCia_tl = '" + Trim(Label15.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(0).Cells(11).Value) + "'
                            ORDER BY  A.CCia_tl, A.Codigo_tl"
                Dim cmd10235 As New SqlCommand(sql10235, conx)
                Rsr205 = cmd10235.ExecuteReader()
                If Rsr205.Read() = True Then
                    DataGridView1.Columns(1).HeaderText = Trim(Rsr205(13))
                    DataGridView1.Columns(2).HeaderText = Trim(Rsr205(14))
                    DataGridView1.Columns(3).HeaderText = Trim(Rsr205(15))
                    DataGridView1.Columns(4).HeaderText = Trim(Rsr205(16))
                    DataGridView1.Columns(5).HeaderText = Trim(Rsr205(17))
                    DataGridView1.Columns(6).HeaderText = Trim(Rsr205(18))
                    DataGridView1.Columns(7).HeaderText = Trim(Rsr205(19))
                    DataGridView1.Columns(8).HeaderText = Trim(Rsr205(20))
                    DataGridView1.Columns(9).HeaderText = Trim(Rsr205(21))
                    DataGridView1.Columns(10).HeaderText = Trim(Rsr205(22))

                If Trim(Rsr205(13)).Length > 0 Then
                    DataGridView1.Rows(0).Cells(1).Value = t1
                    DataGridView1.Rows(1).Cells(1).Value = 0
                    DataGridView1.Columns(1).ReadOnly = False
                Else
                    DataGridView1.Columns(1).ReadOnly = True
                End If
                    If Trim(Rsr205(14)).Length > 0 Then
                    DataGridView1.Rows(0).Cells(2).Value = t2
                    DataGridView1.Rows(1).Cells(2).Value = 0
                    DataGridView1.Columns(2).ReadOnly = False
                Else
                    DataGridView1.Columns(2).ReadOnly = True
                End If
                    If Trim(Rsr205(15)).Length > 0 Then
                    DataGridView1.Rows(0).Cells(3).Value = t3
                    DataGridView1.Rows(1).Cells(3).Value = 0
                    DataGridView1.Columns(3).ReadOnly = True
                    DataGridView1.Columns(3).ReadOnly = False
                Else
                    DataGridView1.Columns(3).ReadOnly = True
                End If
                    If Trim(Rsr205(16)).Length > 0 Then
                    DataGridView1.Rows(0).Cells(4).Value = t4
                    DataGridView1.Rows(1).Cells(4).Value = 0
                    DataGridView1.Columns(4).ReadOnly = False
                Else
                    DataGridView1.Columns(4).ReadOnly = True
                End If
                    If Trim(Rsr205(17)).Length > 0 Then
                    DataGridView1.Rows(0).Cells(5).Value = t5
                    DataGridView1.Rows(1).Cells(5).Value = 0
                    DataGridView1.Columns(5).ReadOnly = False
                Else
                    DataGridView1.Columns(5).ReadOnly = True

                End If
                    If Trim(Rsr205(18)).Length > 0 Then
                    DataGridView1.Rows(0).Cells(6).Value = t6
                    DataGridView1.Rows(1).Cells(6).Value = 0
                    DataGridView1.Columns(6).ReadOnly = False
                Else
                    DataGridView1.Columns(6).ReadOnly = True
                End If
                    If Trim(Rsr205(19)).Length > 0 Then
                    DataGridView1.Rows(0).Cells(7).Value = t7
                    DataGridView1.Rows(1).Cells(7).Value = 0
                    DataGridView1.Columns(7).ReadOnly = False
                Else
                    DataGridView1.Columns(7).ReadOnly = True
                End If
                    If Trim(Rsr205(20)).Length > 0 Then
                    DataGridView1.Rows(0).Cells(8).Value = t8
                    DataGridView1.Rows(1).Cells(8).Value = 0
                    DataGridView1.Columns(8).ReadOnly = False
                Else
                    DataGridView1.Columns(8).ReadOnly = True
                End If
                    If Trim(Rsr205(21)).Length > 0 Then
                    DataGridView1.Rows(0).Cells(9).Value = t9
                    DataGridView1.Rows(1).Cells(9).Value = 0
                    DataGridView1.Columns(9).ReadOnly = False
                Else
                    DataGridView1.Columns(9).ReadOnly = True
                End If
                    If Trim(Rsr205(22)).Length > 0 Then
                    DataGridView1.Rows(0).Cells(10).Value = t10
                    DataGridView1.Rows(1).Cells(10).Value = 0
                    DataGridView1.Columns(10).ReadOnly = False
                Else
                    DataGridView1.Columns(10).ReadOnly = True
                End If

                End If
                Rsr205.Close()

                Dim sql10248 As String = "SELECT distinct fich_3,c.nomb_10 FROM custom_vianny.dbo.Qap3000 q 
            	inner join custom_vianny.dbo.Qag3000 g on q.Empr_3a = g.Empr_3 and q.Ano_3a=g.Ano_3 and q.talm_3a =g.talm_3 and q.ccom_3a =g.ccom_3 and q.ncom_3a = g.ncom_3  
            	left join custom_vianny.dbo.cag1000 c on q.Empr_3a = c.ccia And g.fich_3 = c.fich_10
                where pedido_3a ='" + Trim(TextBox4.Text) + "' and q.talm_3a ='0500'"
                Dim sql1024 As New SqlCommand(sql10248, conx)
                Rsr24 = sql1024.ExecuteReader()
                If Rsr24.Read() = True Then
                    TextBox1.Text = Rsr24(1)
                End If
                Rsr24.Close()
                Dim sql102487 As String = "select isnull(SUM(canti_3q),0) from custom_vianny.dbo.Qaq300cc q WHERE q.pedido_3q = '" + Trim(TextBox4.Text) + "' and q.Empr_3q = '" + Trim(Label15.Text) + "'"
                Dim sql10247 As New SqlCommand(sql102487, conx)
                Rsr29 = sql10247.ExecuteReader()
                If Rsr29.Read() = True Then
                    TextBox8.Text = Rsr29(0)
                End If
            Rsr29.Close()
            Dim suma4, suma5,total, columnCount As Integer
            suma4 = 0
            suma5 = 0
            columnCount = 0
            total = 0
            For Each column As DataGridViewColumn In DataGridView1.Columns
                If Not String.IsNullOrEmpty(column.HeaderText) Then
                    columnCount += 1
                End If
            Next

            For i4 = 1 To columnCount - 2
                suma4 = suma4 + CInt(DataGridView1.Rows(0).Cells(i4).Value)
                suma5 = suma5 + CInt(DataGridView1.Rows(1).Cells(i4).Value)
            Next
            total = suma4 + suma5
            TextBox15.Text = total
        End If
    End Sub

    Private Sub Liqui_Produccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox4.Select()
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If e.ColumnIndex = 1 Then
            Dim CANT1, CANT2, SUMA4, suma5, SUM9, total, columnCount As Integer
            SUMA4 = 0
            CANT2 = CInt(TextBox8.Text)
            CANT1 = CInt(DataGridView1.Rows(1).Cells(1).Value)
            columnCount = 0
            total = 0
            For Each column As DataGridViewColumn In DataGridView1.Columns
                If Not String.IsNullOrEmpty(column.HeaderText) Then
                    columnCount += 1
                End If
            Next


            If CANT1 + CInt(TextBox15.Text) > CANT2 Then
                TextBox16.Text = "LA CANTIDAD INGRESADA ES MAYOR AL PEDIDO"
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = 0
                For i4 = 1 To columnCount - 2
                    SUMA4 = SUMA4 + CInt(DataGridView1.Rows(0).Cells(i4).Value)
                    suma5 = suma5 + CInt(DataGridView1.Rows(1).Cells(i4).Value)
                Next
                total = SUMA4 + suma5
                TextBox15.Text = total
            Else
                For i4 = 1 To columnCount - 2
                    SUMA4 = SUMA4 + CInt(DataGridView1.Rows(0).Cells(i4).Value)
                    suma5 = suma5 + CInt(DataGridView1.Rows(1).Cells(i4).Value)
                Next
                total = SUMA4 + suma5

                SUM9 = CANT2 - total
                TextBox16.Text = "FALTA  " & " " & SUM9 & " " & " PARA LIQUIDAR EL CORTE"
                TextBox15.Text = total
            End If
        End If

        'If e.ColumnIndex = 2 Then
        '    Dim CANT1, CANT2, SUMA4, SUM9 As Integer
        '    SUMA4 = 0
        '    CANT2 = CInt(TextBox7.Text)
        '    CANT1 = CInt(DataGridView1.Rows(e.RowIndex).Cells(2).Value)

        '    If CANT1 + CInt(TextBox15.Text) > CANT2 Then
        '        TextBox16.Text = "LA CANTIDAD INGRESADA ES MAYOR AL PEDIDO"
        '        DataGridView1.Rows(e.RowIndex).Cells(2).Value = 0

        '    Else

        '        For i4 = 1 To 10
        '            SUMA4 = SUMA4 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
        '        Next
        '        SUM9 = CANT2 - SUMA4
        '        TextBox16.Text = "FALTA  " & " " & SUM9 & " " & " PARA LIQUIDAR EL CORTE"
        '        TextBox15.Text = SUMA4
        '    End If
        'End If

        'If e.ColumnIndex = 3 Then
        '    Dim CANT1, CANT2, SUMA4, SUM9 As Integer
        '    SUMA4 = 0
        '    CANT2 = CInt(TextBox7.Text)
        '    CANT1 = CInt(DataGridView1.Rows(e.RowIndex).Cells(3).Value)

        '    If CANT1 + CInt(TextBox15.Text) > CANT2 Then
        '        TextBox16.Text = "LA CANTIDAD INGRESADA ES MAYOR AL PEDIDO"
        '        DataGridView1.Rows(e.RowIndex).Cells(3).Value = 0

        '    Else

        '        For i4 = 1 To 10
        '            SUMA4 = SUMA4 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
        '        Next
        '        SUM9 = CANT2 - SUMA4
        '        TextBox16.Text = "FALTA  " & " " & SUM9 & " " & " PARA LIQUIDAR EL CORTE"
        '        TextBox15.Text = SUMA4
        '    End If
        'End If

        'If e.ColumnIndex = 4 Then
        '    Dim CANT1, CANT2, SUMA4, SUM9 As Integer
        '    SUMA4 = 0
        '    CANT2 = CInt(TextBox7.Text)
        '    CANT1 = CInt(DataGridView1.Rows(e.RowIndex).Cells(4).Value)

        '    If CANT1 + CInt(TextBox15.Text) > CANT2 Then
        '        TextBox16.Text = "LA CANTIDAD INGRESADA ES MAYOR AL PEDIDO"
        '        DataGridView1.Rows(e.RowIndex).Cells(4).Value = 0

        '    Else

        '        For i4 = 1 To 10
        '            SUMA4 = SUMA4 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
        '        Next
        '        SUM9 = CANT2 - SUMA4
        '        TextBox16.Text = "FALTA  " & " " & SUM9 & " " & " PARA LIQUIDAR EL CORTE"
        '        TextBox15.Text = SUMA4
        '    End If

        'End If

        'If e.ColumnIndex = 5 Then
        '    Dim CANT1, CANT2, SUMA4, SUM9 As Integer
        '    SUMA4 = 0
        '    CANT2 = CInt(TextBox7.Text)
        '    CANT1 = CInt(DataGridView1.Rows(e.RowIndex).Cells(5).Value)

        '    If CANT1 + CInt(TextBox15.Text) > CANT2 Then
        '        TextBox16.Text = "LA CANTIDAD INGRESADA ES MAYOR AL PEDIDO"
        '        DataGridView1.Rows(e.RowIndex).Cells(5).Value = 0

        '    Else

        '        For i4 = 1 To 10
        '            SUMA4 = SUMA4 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
        '        Next
        '        SUM9 = CANT2 - SUMA4
        '        TextBox16.Text = "FALTA  " & " " & SUM9 & " " & " PARA LIQUIDAR EL CORTE"
        '        TextBox15.Text = SUMA4
        '    End If
        'End If

        'If e.ColumnIndex = 6 Then
        '    Dim CANT1, CANT2, SUMA4, SUM9 As Integer
        '    SUMA4 = 0
        '    CANT2 = CInt(TextBox7.Text)
        '    CANT1 = CInt(DataGridView1.Rows(e.RowIndex).Cells(6).Value)

        '    If CANT1 + CInt(TextBox15.Text) > CANT2 Then
        '        TextBox16.Text = "LA CANTIDAD INGRESADA ES MAYOR AL PEDIDO"
        '        DataGridView1.Rows(e.RowIndex).Cells(6).Value = 0

        '    Else

        '        For i4 = 1 To 10
        '            SUMA4 = SUMA4 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
        '        Next
        '        SUM9 = CANT2 - SUMA4
        '        TextBox16.Text = "FALTA  " & " " & SUM9 & " " & " PARA LIQUIDAR EL CORTE"
        '        TextBox15.Text = SUMA4
        '    End If
        'End If

        'If e.ColumnIndex = 7 Then
        '    Dim CANT1, CANT2, SUMA4, SUM9 As Integer
        '    SUMA4 = 0
        '    CANT2 = CInt(TextBox7.Text)
        '    CANT1 = CInt(DataGridView1.Rows(e.RowIndex).Cells(7).Value)

        '    If CANT1 + CInt(TextBox15.Text) > CANT2 Then
        '        TextBox16.Text = "LA CANTIDAD INGRESADA ES MAYOR AL PEDIDO"
        '        DataGridView1.Rows(e.RowIndex).Cells(7).Value = 0

        '    Else

        '        For i4 = 1 To 10
        '            SUMA4 = SUMA4 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
        '        Next
        '        SUM9 = CANT2 - SUMA4
        '        TextBox16.Text = "FALTA  " & " " & SUM9 & " " & " PARA LIQUIDAR EL CORTE"
        '        TextBox15.Text = SUMA4
        '    End If
        'End If

        'If e.ColumnIndex = 8 Then
        '    Dim CANT1, CANT2, SUMA4, SUM9 As Integer
        '    SUMA4 = 0
        '    CANT2 = CInt(TextBox7.Text)
        '    CANT1 = CInt(DataGridView1.Rows(e.RowIndex).Cells(8).Value)

        '    If CANT1 + CInt(TextBox15.Text) > CANT2 Then
        '        TextBox16.Text = "LA CANTIDAD INGRESADA ES MAYOR AL PEDIDO"
        '        DataGridView1.Rows(e.RowIndex).Cells(8).Value = 0

        '    Else

        '        For i4 = 1 To 10
        '            SUMA4 = SUMA4 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
        '        Next
        '        SUM9 = CANT2 - SUMA4
        '        TextBox16.Text = "FALTA  " & " " & SUM9 & " " & " PARA LIQUIDAR EL CORTE"
        '        TextBox15.Text = SUMA4
        '    End If
        'End If

        'If e.ColumnIndex = 9 Then
        '    Dim CANT1, CANT2, SUMA4, SUM9 As Integer
        '    SUMA4 = 0
        '    CANT2 = CInt(TextBox7.Text)
        '    CANT1 = CInt(DataGridView1.Rows(e.RowIndex).Cells(9).Value)

        '    If CANT1 + CInt(TextBox15.Text) > CANT2 Then
        '        TextBox16.Text = "LA CANTIDAD INGRESADA ES MAYOR AL PEDIDO"
        '        DataGridView1.Rows(e.RowIndex).Cells(9).Value = 0

        '    Else

        '        For i4 = 1 To 10
        '            SUMA4 = SUMA4 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
        '        Next
        '        SUM9 = CANT2 - SUMA4
        '        TextBox16.Text = "FALTA  " & " " & SUM9 & " " & " PARA LIQUIDAR EL CORTE"
        '        TextBox15.Text = SUMA4
        '    End If
        'End If

        'If e.ColumnIndex = 10 Then
        '    Dim CANT1, CANT2, SUMA4, SUM9 As Integer
        '    SUMA4 = 0
        '    CANT2 = CInt(TextBox7.Text)
        '    CANT1 = CInt(DataGridView1.Rows(e.RowIndex).Cells(10).Value)

        '    If CANT1 + CInt(TextBox15.Text) > CANT2 Then
        '        TextBox16.Text = "LA CANTIDAD INGRESADA ES MAYOR AL PEDIDO"
        '        DataGridView1.Rows(e.RowIndex).Cells(9).Value = 0

        '    Else

        '        For i4 = 1 To 10
        '            SUMA4 = SUMA4 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
        '        Next
        '        SUM9 = CANT2 - SUMA4
        '        TextBox16.Text = "FALTA  " & " " & SUM9 & " " & " PARA LIQUIDAR EL CORTE"
        '        TextBox15.Text = SUMA4
        '    End If
        'End If
    End Sub
End Class