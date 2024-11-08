Imports System.Data.SqlClient
Public Class ORDEN_CORTE
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
    Dim Rsr21, Rsr2, Rsr222 As SqlDataReader
    Private Sub ORDEN_CORTE_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim sql1021 As String = "SELECT OP,tela,G.nomb_17,cxplineal FROM custom_vianny.dbo.Consumo_Op_DET C LEFT JOIN custom_vianny.dbo.cag1700 g ON  c.ccia = g.ccia and  c.tela = g.linea_17 WHERE OP='" + Trim(Label8.Text) + "'"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr21 = cmd1021.ExecuteReader()
        Dim i51 As Integer
        i51 = 0

        While Rsr21.Read()
            DataGridView2.Rows.Add()
            DataGridView2.Rows(i51).Cells(0).Value = Rsr21(1)
            DataGridView2.Rows(i51).Cells(1).Value = Rsr21(0)
            DataGridView2.Rows(i51).Cells(2).Value = Rsr21(2)
            DataGridView2.Rows(i51).Cells(3).Value = Rsr21(3)
            DataGridView2.Rows(i51).Cells(4).Value = 0
            i51 = i51 + 1
        End While
        Rsr21.Close()
        If Label5.Text <> 1 Then
            ULTIMO()
        End If

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
        TextBox1.Select()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        Dim cmd15 As New SqlCommand("insert into orden_corte (id_corte,ruc,proveedor,po,vendedor,cliente,modelista,ruta,fecha,ccia,periodo) values (@id_corte,@ruc,@proveedor,@po,@vendedor,@cliente,@modelista,@ruta,@fecha,@ccia,@periodo)", conx)
        cmd15.Parameters.AddWithValue("@id_corte", Trim(TextBox1.Text))
        cmd15.Parameters.AddWithValue("@ruc", Trim(TextBox3.Text))
        cmd15.Parameters.AddWithValue("@proveedor", Trim(TextBox5.Text))
        cmd15.Parameters.AddWithValue("@po", Trim(TextBox2.Text))
        cmd15.Parameters.AddWithValue("@vendedor", Trim(TextBox6.Text))
        cmd15.Parameters.AddWithValue("@cliente", Trim(TextBox4.Text))
        cmd15.Parameters.AddWithValue("@modelista", Trim(TextBox7.Text))
        cmd15.Parameters.AddWithValue("@ruta", Trim(TextBox8.Text))
        cmd15.Parameters.AddWithValue("@fecha", DateTimePicker1.Value)
        cmd15.Parameters.AddWithValue("@ccia", Trim(Label12.Text))
        cmd15.Parameters.AddWithValue("@periodo", Trim(Label8.Text))
        cmd15.ExecuteNonQuery()

        Dim j, j1 As Integer
        j = DataGridView1.Rows.Count
        For i = 0 To j - 1
            Dim cmd16 As New SqlCommand("insert into orden_corte_po (op,producto,estilo,cantidad,consumo,total,id_corte,fecha,ccia,periodo) values (@op,@producto,@estilo,@cantidad,@consumo,@total,@id_corte,@fecha,@ccia,@periodo)", conx)
            cmd16.Parameters.AddWithValue("@op", Trim(DataGridView1.Rows(i).Cells(0).Value))
            cmd16.Parameters.AddWithValue("@producto", Trim(DataGridView1.Rows(i).Cells(1).Value))
            cmd16.Parameters.AddWithValue("@estilo", Trim(DataGridView1.Rows(i).Cells(2).Value))
            cmd16.Parameters.AddWithValue("@cantidad", Trim(DataGridView1.Rows(i).Cells(3).Value))
            cmd16.Parameters.AddWithValue("@consumo", Trim(DataGridView1.Rows(i).Cells(4).Value))
            cmd16.Parameters.AddWithValue("@total", Trim(DataGridView1.Rows(i).Cells(5).Value))
            cmd16.Parameters.AddWithValue("@id_corte", Trim(TextBox1.Text))
            cmd16.Parameters.AddWithValue("@fecha", DateTimePicker1.Value)
            cmd16.Parameters.AddWithValue("@ccia", Trim(Label12.Text))
            cmd16.Parameters.AddWithValue("@periodo", Trim(Label8.Text))
            cmd16.ExecuteNonQuery()

            Dim cmd18 As New SqlCommand("update custom_vianny.dbo.qag0300 set hruta12_3 =1 where ncom_3 ='" + Trim(DataGridView1.Rows(i).Cells(0).Value) + "' and ccia ='" + Trim(Label12.Text) + "'", conx)
            cmd18.Parameters.AddWithValue("@op", Trim(DataGridView1.Rows(i).Cells(0).Value))
            cmd18.Parameters.AddWithValue("@producto", Trim(DataGridView1.Rows(i).Cells(1).Value))
            cmd18.ExecuteNonQuery()
        Next

        j1 = DataGridView2.Rows.Count
        For i1 = 0 To j1 - 1
            Dim cmd17 As New SqlCommand("insert into orden_corte_tela (articulo,tela,kg,id_corte,fecha,ccia,periodo) values (@articulo,@tela,@kg,@id_corte,@fecha,@ccia,@periodo)", conx)
            cmd17.Parameters.AddWithValue("@articulo", Trim(DataGridView2.Rows(i1).Cells(0).Value))
            cmd17.Parameters.AddWithValue("@tela", Trim(DataGridView2.Rows(i1).Cells(1).Value))
            cmd17.Parameters.AddWithValue("@kg", Trim(DataGridView2.Rows(i1).Cells(2).Value))
            cmd17.Parameters.AddWithValue("@id_corte", Trim(TextBox1.Text))
            cmd17.Parameters.AddWithValue("@fecha", DateTimePicker1.Value)
            cmd17.Parameters.AddWithValue("@ccia", Trim(Label12.Text))
            cmd17.Parameters.AddWithValue("@periodo", Trim(Label8.Text))
            cmd17.ExecuteNonQuery()

        Next
        MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
        MsgBox("SE GENERO EL REPORTE DE ORDEN DE CORTE")
        Rp_corte.TextBox1.Text = Label12.Text
        Rp_corte.TextBox2.Text = TextBox1.Text
        Rp_corte.ShowDialog()

        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA IMPRIMIR LA GUIA DE REMISION? SI ELIGE NO SE GENERARA UNA NOTA DE SALIDA", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            guia_talleres.Label23.Text = "0300"
            guia_talleres.Label25.Text = "CORTE"
            guia_talleres.TextBox24.Text = "0300"
            guia_talleres.TextBox23.Text = "CORTE"
            guia_talleres.Label30.Text = Label12.Text
            guia_talleres.Label29.Text = Label8.Text
            Dim y As Integer
            y = DataGridView1.Rows.Count

            For i = 0 To y - 1
                abrir()
                Dim VAL As Integer
                guia_talleres.TextBox25.Text = DataGridView1.Rows(i).Cells(0).Value
                Select Case Trim(guia_talleres.TextBox25.Text).Length
                    Case "1" : guia_talleres.TextBox25.Text = "01" & "0000000" & guia_talleres.TextBox25.Text
                    Case "2" : guia_talleres.TextBox25.Text = "01" & "000000" & guia_talleres.TextBox25.Text
                    Case "3" : guia_talleres.TextBox25.Text = "01" & "00000" & guia_talleres.TextBox25.Text
                    Case "4" : guia_talleres.TextBox25.Text = "01" & "0000" & guia_talleres.TextBox25.Text
                    Case "5" : guia_talleres.TextBox25.Text = "01" & "000" & guia_talleres.TextBox25.Text
                    Case "6" : guia_talleres.TextBox25.Text = "01" & "00" & guia_talleres.TextBox25.Text
                    Case "7" : guia_talleres.TextBox25.Text = "01" & "0" & guia_talleres.TextBox25.Text
                    Case "8" : guia_talleres.TextBox25.Text = "01" & guia_talleres.TextBox25.Text

                End Select
                Dim sql102 As String = "select ncom_3,linea_3,a.nomb_17, a.unid_17 from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + Trim(guia_talleres.TextBox25.Text) + "' and q.ccia ='" + Trim(guia_talleres.Label30.Text) + "'
"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr222 = cmd102.ExecuteReader()
                Dim i5 As Integer
                i5 = guia_talleres.DataGridView1.Rows.Count

                If Rsr222.Read() = True Then
                    guia_talleres.DataGridView1.Rows.Add()
                    guia_talleres.DataGridView1.Rows(i5).Cells(1).Value = Rsr222(0)
                    guia_talleres.DataGridView1.Rows(i5).Cells(2).Value = Rsr222(1)
                    guia_talleres.DataGridView1.Rows(i5).Cells(3).Value = Rsr222(2)
                    guia_talleres.DataGridView1.Rows(i5).Cells(5).Value = Rsr222(3)
                    guia_talleres.DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                    VAL = guia_talleres.DataGridView1.Rows(i5).Cells(0).Value
                    Select Case VAL.ToString.Length
                        Case "1" : guia_talleres.DataGridView1.Rows(i5).Cells(0).Value = "00" & "" & VAL
                        Case "2" : guia_talleres.DataGridView1.Rows(i5).Cells(0).Value = "0" & "" & VAL
                        Case "3" : guia_talleres.DataGridView1.Rows(i5).Cells(0).Value = VAL
                    End Select
                    guia_talleres.TextBox25.Text = ""

                    'DataGridView1.Rows(i5).Cells(4).Selected = True
                    guia_talleres.DataGridView1.CurrentCell = guia_talleres.DataGridView1.Rows(i5).Cells(4)
                    guia_talleres.DataGridView1.BeginEdit(True)
                    Rsr222.Close()
                    If guia_talleres.Label23.Text = "0400" Then
                        abrir()
                        guia_talleres.llenar_combo_box3()
                    End If
                Else
                    MsgBox("LA OP INGRESADA NO EXISTE")
                End If
                Rsr222.Close()

            Next
            guia_talleres.Label33.Text = 1
            guia_talleres.ShowDialog()
        Else
            MsgBox("SE GENERARA UN NOTA DE SALIDA")
            nsa_produccion.Close()
            nsa_produccion.TextBox1.Text = "0300"
            nsa_produccion.TextBox13.Text = "0300"
            nsa_produccion.TextBox14.Text = "0321"
            nsa_produccion.TextBox16.Text = "CORTE"
            nsa_produccion.TextBox11.Text = "CORTE"
            nsa_produccion.TextBox2.Text = "AREA CORTE"
            nsa_produccion.Label11.Text = Label8.Text
            nsa_produccion.Label4.Text = Label12.Text

            Dim y1 As Integer
            y1 = DataGridView1.Rows.Count

            For i1 = 0 To y1 - 1

                abrir()

                nsa_produccion.TextBox8.Text = TextBox3.Text
                nsa_produccion.TextBox10.Text = TextBox5.Text
                nsa_produccion.TextBox15.Text = DataGridView1.Rows(i1).Cells(0).Value
                Dim VAL As Integer
                Select Case Trim(nsa_produccion.TextBox15.Text).Length
                    Case "1" : nsa_produccion.TextBox15.Text = "01" & "0000000" & nsa_produccion.TextBox15.Text
                    Case "2" : nsa_produccion.TextBox15.Text = "01" & "000000" & nsa_produccion.TextBox15.Text
                    Case "3" : nsa_produccion.TextBox15.Text = "01" & "00000" & nsa_produccion.TextBox15.Text
                    Case "4" : nsa_produccion.TextBox15.Text = "01" & "0000" & nsa_produccion.TextBox15.Text
                    Case "5" : nsa_produccion.TextBox15.Text = "01" & "000" & nsa_produccion.TextBox15.Text
                    Case "6" : nsa_produccion.TextBox15.Text = "01" & "00" & nsa_produccion.TextBox15.Text
                    Case "7" : nsa_produccion.TextBox15.Text = "01" & "0" & nsa_produccion.TextBox15.Text
                    Case "8" : nsa_produccion.TextBox15.Text = "01" & nsa_produccion.TextBox15.Text

                End Select
                Dim sql102 As String = "select ncom_3,linea_3,a.nomb_17, a.unid_17 from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + Trim(nsa_produccion.TextBox15.Text) + "' and q.ccia ='" + Trim(nsa_produccion.Label4.Text) + "'
"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()
                Dim i5 As Integer
                i5 = nsa_produccion.DataGridView1.Rows.Count

                If Rsr2.Read() = True Then
                    nsa_produccion.DataGridView1.Rows.Add()
                    nsa_produccion.DataGridView1.Rows(i5).Cells(1).Value = Rsr2(0)
                    nsa_produccion.DataGridView1.Rows(i5).Cells(2).Value = Rsr2(1)
                    nsa_produccion.DataGridView1.Rows(i5).Cells(3).Value = Rsr2(2)
                    nsa_produccion.DataGridView1.Rows(i5).Cells(5).Value = Rsr2(3)
                    nsa_produccion.DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                    VAL = nsa_produccion.DataGridView1.Rows(i5).Cells(0).Value
                    Select Case VAL.ToString.Length
                        Case "1" : nsa_produccion.DataGridView1.Rows(i5).Cells(0).Value = "00" & "" & VAL
                        Case "2" : nsa_produccion.DataGridView1.Rows(i5).Cells(0).Value = "0" & "" & VAL
                        Case "3" : nsa_produccion.DataGridView1.Rows(i5).Cells(0).Value = VAL
                    End Select
                    nsa_produccion.TextBox15.Text = ""

                    'DataGridView1.Rows(i5).Cells(4).Selected = True
                    nsa_produccion.DataGridView1.CurrentCell = nsa_produccion.DataGridView1.Rows(i5).Cells(4)
                    nsa_produccion.DataGridView1.BeginEdit(True)
                    Rsr2.Close()
                    If TextBox1.Text = "0400" Then
                        abrir()
                        nsa_produccion.llenar_combo_box3()
                    End If

                Else
                    MsgBox("LA OP INGRESADA NO EXISTE")
                End If
                Rsr2.Close()


            Next

            nsa_produccion.Label16.Text = 1
            nsa_produccion.ShowDialog()
        End If


        Me.Close()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Rp_corte.TextBox1.Text = Label12.Text
        Rp_corte.TextBox2.Text = TextBox1.Text
        Rp_corte.Show()
    End Sub

    Sub ULTIMO()
        abrir()
        Dim sql102 As String = "SELECT TOP 1 id_corte  FROM orden_corte  order by  id_corte desc"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        Dim i5 As Integer
        i5 = 0
        If Rsr2.Read() = True Then
            TextBox1.Text = Rsr2(0) + 1
        Else
            TextBox1.Text = 1
        End If
        Rsr2.Close()
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.F1 Then
            det_ns_Prodc.Label2.Text = "0300"
            det_ns_Prodc.Label3.Text = 8
            det_ns_Prodc.Label4.Text = "1"
            det_ns_Prodc.Label5.Text = Trim(nsa_produccion.Label4.Text)
            det_ns_Prodc.Show()
        End If
    End Sub
End Class