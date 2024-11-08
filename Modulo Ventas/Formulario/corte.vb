Imports System.ComponentModel
Imports System.Data.SqlClient
Public Class corte
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
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
    Dim Rs1, RS2, Rs10, RS3, RS4, RS200 As SqlDataReader
    Dim ju2, ju3 As New DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
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
            Dim sql10 As String = "	SELECT COUNT(ncom_3) FROM  custom_vianny.dbo.qag0300  WHERE ncom_3 ='" + Trim(TextBox1.Text) + "'  and ccia ='" + Trim(Label8.Text) + "' AND flag_3=1 "
            Dim cmd10 As New SqlCommand(sql10, conx)
            Rs10 = cmd10.ExecuteReader()
            If Rs10.Read() = True Then
                If Rs10(0) = 0 Then
                    MsgBox("LA OP INGRESADA NO EXISTE")
                Else
                    TextBox2.Enabled = True
                    TextBox5.Enabled = True
                    TextBox6.Enabled = True
                    DataGridView1.Enabled = True
                    TextBox2.Select()

                End If



            End If
            Rs10.Close()
        End If


    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        CALCULO()

    End Sub
    Dim da55 As New DataTable
    Dim da1 As New DataTable
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Clear()
            'DataGridView2.Rows.Clear()
            DataGridView2.Columns.Clear()
            da55.Clear()
            da1.Clear()
            'Dim u As Integer
            'u = DataGridView2.Columns.Count
            'If u > 1 Then
            '    For X = 1 To DataGridView2.Columns.Count - 1

            '        DataGridView2.Columns.Remove(DataGridView2.Columns(X))

            '    Next
            'End If
            TextBox5.Text = 0
            TextBox6.Text = 0
            TextBox13.Text = 0
            abrir()
            Dim sql1 As String = "SELECT descri_3,tela1_3,telaprinc_3,telasec_3, program_3,cantp_3,b.color_16b,CONSUMO,CANT_CONSU,G.unid_17 FROM custom_vianny.dbo.qag0300 Q 
inner join custom_vianny.dbo.qag160b b on q.ccia = b.ccia  and q.ncom_3 = b.ncom_16b
LEFT JOIN custom_vianny.dbo.cag1700 G ON Q.tela1_3 = G.linea_17 AND Q.ccia=G.ccia 
WHERE ncom_3 ='" + Trim(TextBox1.Text) + "' AND q.ccia='" + Label8.Text + "'"
            Dim cmd1 As New SqlCommand(sql1, conx)
            Rs1 = cmd1.ExecuteReader()
            If Rs1.Read() = True Then
                TextBox7.Text = Rs1(0)
                TextBox3.Text = Rs1(2)
                TextBox4.Text = Rs1(5)

                TextBox10.Text = Rs1(6)
                If Rs1(7) Is DBNull.Value Then
                    TextBox9.Text = "0.00"
                Else
                    TextBox9.Text = Rs1(7)
                End If
                If Rs1(8) Is DBNull.Value Then
                    TextBox11.Text = "0.00"
                Else
                    TextBox11.Text = Rs1(8)
                End If

                TextBox12.Text = Rs1(4)
                If Rs1(9) Is DBNull.Value Then
                    TextBox8.Text = ""
                    Label13.Text = ""
                Else
                    TextBox8.Text = Rs1(9)
                    Label13.Text = Rs1(9)
                End If


                Rs1.Close()

                Dim KL As New fop
                Dim OL As New vop

                OL.gcia = Label8.Text
                OL.gncom_3 = TextBox1.Text

                abrir()

                Dim sql2 As String = "SELECT COUNT(cant_16d) FROM custom_vianny.dbo.Qag16dv where  ccia_16d = '" + Trim(Label8.Text) + "' And ncom_16d = '" + TextBox1.Text + "' "
                Dim cmd2 As New SqlCommand(sql2, conx)
                RS2 = cmd2.ExecuteReader()
                RS2.Read()

                If RS2(0) <> 0 Then

                    ju2 = KL.PADRON_TALLA1(OL)
                    DataGridView6.DataSource = ju2

                    ju3 = KL.PADRON_TALLA(OL)
                    DataGridView5.DataSource = ju3

                    Dim K As Integer
                    K = DataGridView5.Columns.Count
                    DataGridView2.Columns.Add("ESTADO", "ESTADO")
                    DataGridView2.Rows.Add(3)
                    DataGridView2.Rows(0).Cells(0).Value = "SOLICITADO"
                    DataGridView2.Rows(1).Cells(0).Value = "PROGRAMADO"
                    DataGridView2.Rows(2).Cells(0).Value = "CORTADO"
                    For I = 0 To K - 1
                        DataGridView2.Columns.Add(DataGridView5.Columns(I).HeaderText, DataGridView5.Columns(I).HeaderText)
                        DataGridView2.Rows(0).Cells(I + 1).Value = DataGridView5.Rows(0).Cells(I).Value
                        DataGridView2.Rows(1).Cells(I + 1).Value = DataGridView6.Rows(0).Cells(I).Value
                    Next
                    Label15.Text = RS2(0)
                Else

                End If
                DataGridView5.DataSource = Nothing
                DataGridView6.DataSource = Nothing
                RS2.Close()

                'DataGridView2.Rows(2).Cells(0).Value = "CORTADO"
                abrir()
                Dim cmd55 As New SqlDataAdapter("select talla ,SUM(cantidad) from detalle_TENDIDO_CORTE where op ='" + Trim(TextBox1.Text) + "' and CORTE ='" + Trim(TextBox2.Text) + "' group by talla ", conx)
                cmd55.Fill(da55)
                If da55.Rows.Count > 0 Then
                    DataGridView7.DataSource = da55
                Else
                    DataGridView7.DataSource = Nothing
                End If

                Dim f, f1 As Integer
                f = DataGridView7.Rows.Count
                f1 = DataGridView2.Columns.Count

                For i3 = 1 To f1 - 1

                    For i1 = 0 To f - 1
                        'MsgBox(DataGridView2.Rows(2).Cells(i1 + 1).Value)
                        'MsgBox(DataGridView2.Columns(i3).HeaderText.ToString.Trim)
                        'MsgBox(DataGridView7.Rows(i1).Cells(0).Value.ToString.Trim)
                        If DataGridView2.Columns(i3).HeaderText.ToString.Trim = DataGridView7.Rows(i1).Cells(0).Value.ToString.Trim Then
                            DataGridView2.Rows(2).Cells(i1 + 1).Value = DataGridView7.Rows(i1).Cells(1).Value
                        End If

                    Next
                Next
                DataGridView7.DataSource = Nothing
            End If


            '            DataGridView1.Rows.Add(2)
            '            Dim u As Integer
            '            u = DataGridView1.Rows.Count
            '            For i = 0 To u - 1
            '                DataGridView1.Rows(i).Cells(0).Value = TextBox1.Text
            '                DataGridView1.Rows(i).Cells(1).Value = i + 1
            '            Next
            '            TextBox2.Text = ""
            Dim sql3 As String = "select COUNT(OP)  from  TENDIDO_CORTE where OP='" + TextBox1.Text + "' and CORTE ='" + TextBox2.Text + "'"
            Dim cmd3 As New SqlCommand(sql3, conx)
            RS3 = cmd3.ExecuteReader()
            RS3.Read()
            If RS3(0) > 0 Then
                RS3.Close()
                Dim sql4 As String = "select CONSUMOCORTE,CORTADO_CORTE from  TENDIDO_CORTE where OP='" + TextBox1.Text + "' and CORTE ='" + TextBox2.Text + "'"
                Dim cmd4 As New SqlCommand(sql4, conx)
                RS4 = cmd4.ExecuteReader()
                RS4.Read()
                TextBox5.Text = RS4(0)
                TextBox6.Text = RS4(1)
                RS4.Close()
                abrir()
                Dim cmd12 As New SqlDataAdapter("select talla,cantidad,INICIO,FIN from detalle_TENDIDO_CORTE  where OP='" + TextBox1.Text + "' and CORTE ='" + TextBox2.Text + "'", conx)



                cmd12.Fill(da1)
                DataGridView4.DataSource = da1
                Dim ui As Integer
                ui = DataGridView4.Rows.Count
                DataGridView1.Rows.Add(ui)
                For i = 0 To ui - 1
                    DataGridView1.Rows(i).Cells(0).Value = i + 1
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView4.Rows(i).Cells(0).Value
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView4.Rows(i).Cells(1).Value
                    DataGridView1.Rows(i).Cells(3).Value = DataGridView4.Rows(i).Cells(2).Value
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView4.Rows(i).Cells(3).Value
                Next

            Else
                RS3.Close()
                Dim cmd55 As New SqlDataAdapter("SELECT talla_16c FROM custom_vianny.dbo.qag160d d left join custom_vianny.dbo.qag160c c on d.ccia=c.ccia and d.ncom_16d=c.ncom_16c and c.correl_16c=d.correl_16d
                Where  d.ccia = '" + Trim(Label8.Text) + "' And ncom_16d = '" + TextBox1.Text + "' And color_16d = '01'", conx)

                Dim da1 As New DataTable
                Dim SUMA, SUMA1 As Integer
                cmd55.Fill(da1)
                If Label15.Text > 0 Then
                    DataGridView3.DataSource = da1
                    SUMA = 1
                    SUMA1 = 1
                    Dim y As Integer
                    y = DataGridView3.Rows.Count
                    DataGridView1.Rows.Add(y)
                    For i1 = 0 To y - 1
                        DataGridView1.Rows(i1).Cells(0).Value = i1 + 1
                        DataGridView1.Rows(i1).Cells(1).Value = DataGridView3.Rows(i1).Cells(0).Value
                        DataGridView1.Rows(i1).Cells(2).Value = DataGridView2.Rows(1).Cells(i1 + 1).Value
                        DataGridView1.Rows(i1).Cells(3).Value = SUMA1
                        'DataGridView1.Rows(i1).Cells(4).Value = DataGridView1.Rows(i1).Cells(2).Value + DataGridView1.Rows(i1).Cells(3).Value - 1
                        'SUMA = SUMA + DataGridView1.Rows(i1).Cells(1).Value
                        SUMA1 = 1 + DataGridView1.Rows(i1).Cells(4).Value
                    Next
                Else
                    DataGridView3.DataSource = ""
                    DataGridView1.Rows.Clear()
                End If

            End If

            TextBox5.Select()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim YI As Integer
        YI = DataGridView1.Rows.Count
        Dim PP As Double
        PP = 0
        For I = 0 To YI - 1
            PP = PP + DataGridView1.Rows(I).Cells(3).Value
        Next

        If Me.ValidateChildren() And TextBox5.Text = String.Empty Or TextBox6.Text = String.Empty Or TextBox2.Text = String.Empty Or PP = 0 Or TextBox1.Text = String.Empty Then
            MsgBox("FALTAN INGREAR ALGUNO CAMPOS OBLIGATORIOS O LA CANTIDAD CORTADA POR TALLA")
        Else
            abrir()

            Dim agregar As String = "delete from TENDIDO_CORTE where OP='" + Trim(TextBox1.Text) + "' and CORTE ='" + Trim(TextBox2.Text) + "'"
            ELIMINAR(agregar)
            Dim agregar1 As String = "delete from detalle_TENDIDO_CORTE where OP='" + Trim(TextBox1.Text) + "' and CORTE ='" + Trim(TextBox2.Text) + "'"
            ELIMINAR(agregar1)
            Dim cmd20 As New SqlCommand("insert into TENDIDO_CORTE(OP,CORTE,CONSUMO_UDP,CANT_PRODUCCION,TELA,UND,CANT_TELA,CONSUMOUDP,CONSUMOCORTE,CORTADO_CORTE,fecha,PRENDASXCAJA) 
values(@OP,@CORTE,@CONSUMO_UDP,@CANT_PRODUCCION,@TELA,@UND,@CANT_TELA,@CONSUMOUDP,@CONSUMOCORTE,@CORTADO_CORTE,@fecha,@PRENDASXCAJA)", conx)
            cmd20.Parameters.AddWithValue("@OP", Trim(TextBox1.Text))
            cmd20.Parameters.AddWithValue("@CORTE", Trim(TextBox2.Text))
            cmd20.Parameters.AddWithValue("@CONSUMO_UDP", TextBox9.Text)
            cmd20.Parameters.AddWithValue("@CANT_PRODUCCION", TextBox4.Text)
            cmd20.Parameters.AddWithValue("@TELA", Trim(TextBox3.Text))
            cmd20.Parameters.AddWithValue("@UND", Trim(TextBox8.Text))
            cmd20.Parameters.AddWithValue("@CANT_TELA", TextBox11.Text)
            cmd20.Parameters.AddWithValue("@CONSUMOUDP", TextBox9.Text)
            cmd20.Parameters.AddWithValue("@CONSUMOCORTE", TextBox5.Text)
            cmd20.Parameters.AddWithValue("@CORTADO_CORTE", TextBox6.Text)
            cmd20.Parameters.AddWithValue("@fecha", DateTimePicker1.Value)
            cmd20.Parameters.AddWithValue("@PRENDASXCAJA", TextBox13.Text)
            cmd20.ExecuteNonQuery()

            Dim f As Integer
            f = DataGridView1.Rows.Count
            For i = 0 To f - 1
                Dim cmd21 As New SqlCommand("insert into detalle_TENDIDO_CORTE(OP,CORTE,talla,cantidad,INICIO,FIN) values(@OP,@CORTE,@talla,@cantidad,@INICIO,@FIN)", conx)
                cmd21.Parameters.AddWithValue("@OP", Trim(TextBox1.Text))
                cmd21.Parameters.AddWithValue("@CORTE", Trim(TextBox2.Text))
                cmd21.Parameters.AddWithValue("@talla", Trim(DataGridView1.Rows(i).Cells(1).Value))
                cmd21.Parameters.AddWithValue("@cantidad", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd21.Parameters.AddWithValue("@INICIO", Trim(DataGridView1.Rows(i).Cells(3).Value))
                cmd21.Parameters.AddWithValue("@FIN", Trim(DataGridView1.Rows(i).Cells(4).Value))
                cmd21.ExecuteNonQuery()
            Next

        End If


        MsgBox("LA INFORMACION SE REGISTRO CORRECTAMENTE")
        Dim BH As New Rpt_corte
        BH.TextBox1.Text = TextBox1.Text
        BH.TextBox2.Text = TextBox2.Text
        BH.Show()
        Label15.Text = 0
        LIMPIAR()
    End Sub
    Sub LIMPIAR()
        TextBox9.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox11.Text = ""
        TextBox10.Text = ""
        TextBox4.Text = ""
        TextBox8.Text = ""
        TextBox3.Text = ""
        TextBox12.Text = ""
        TextBox7.Text = ""
        TextBox2.Text = ""
        TextBox1.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.DataSource = ""
        DataGridView3.DataSource = ""
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
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Sub CALCULO()
        Dim t As Integer
        t = DataGridView1.Rows.Count
        Dim va As Double

        va = 0
        For O = 0 To t - 1
            va = va + DataGridView1.Rows(O).Cells(1).Value
        Next

        TextBox6.Text = va

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub corte_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown

    End Sub

    Private Sub TextBox5_Validating(sender As Object, e As CancelEventArgs) Handles TextBox5.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL CONSUMO  CORTE")
        End If
    End Sub

    Private Sub TextBox6_Validating(sender As Object, e As CancelEventArgs) Handles TextBox6.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR CANTIDAD PRENDAS CORTADAS")
        End If
    End Sub

    Private Sub TextBox2_Validating(sender As Object, e As CancelEventArgs) Handles TextBox2.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAREL NUMERO DE CORTE")
        End If
    End Sub
    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox2.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAREL LA OP")
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        NumConFrac(TextBox2, e)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Rpt_corte.TextBox1.Text = TextBox1.Text
        Rpt_corte.TextBox2.Text = TextBox2.Text
        Rpt_corte.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DataGridView1.Rows.Add()
        Dim O As Integer
        O = DataGridView1.Rows.Count
        DataGridView1.Rows(O - 1).Cells(0).Value = O
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        NumConFrac(TextBox1, e)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        LIMPIAR()
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

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        NumConFrac1(TextBox5, e)
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        NumConFrac1(TextBox6, e)
    End Sub
End Class