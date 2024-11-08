Imports System.Data.SqlClient
Public Class Generara_ingreso
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
    Dim da As New DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.DataSource = ""
            da.Clear()
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
            Dim cmd As New SqlDataAdapter("select PO,linea_17 AS CODIGO,c.nomb_17 AS PRODUCTO,rollos AS ROLLOS,'RLL' AS UM,peso_total AS PESO,'KGS' AS UM,partida AS PARTIDA,'' AS ITEMS,igsa from rama  r
left join  custom_vianny.dbo.Ranp300 M on r.partida =m.regis_3p
left join custom_vianny.dbo.cag1700 c on m.linea_3p = c.linea_17 and c.ccia = '01'
WHERE codigo ='" + Trim(TextBox1.Text) + "'
", conx)



            cmd.Fill(da)
            If da.Rows.Count <> 0 Then
                DataGridView1.DataSource = da
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(2).Width = 130
                DataGridView1.Columns(3).Width = 360
                DataGridView1.Columns(4).Width = 60
                DataGridView1.Columns(5).Width = 50
                DataGridView1.Columns(6).Width = 85
                DataGridView1.Columns(7).Width = 50
                DataGridView1.Columns(8).Width = 80
                DataGridView1.Columns(8).Width = 60
                DataGridView1.Columns(10).Visible = False
                Dim J As Integer
                J = DataGridView1.Rows.Count
                Dim KP As Integer
                KP = 0
                For I = 0 To J - 1
                    KP = KP + 1
                    DataGridView1.Rows(I).Cells(9).Value = KP
                    If Trim(DataGridView1.Rows(I).Cells(10).Value) = 1 Then
                        DataGridView1.Rows(I).Cells(0).Value = True
                        DataGridView1.Rows(I).ReadOnly = True
                        DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.DarkKhaki
                    End If

                Next
            Else
                da.Clear()
                DataGridView1.DataSource = da
            End If


        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Dim J As Integer
            J = DataGridView1.Rows.Count

            For I = 0 To J - 1
                DataGridView1.Rows(I).Cells(0).Value = True
            Next
        Else
            Dim J As Integer
            J = DataGridView1.Rows.Count

            For I = 0 To J - 1
                DataGridView1.Rows(I).Cells(0).Value = False
            Next
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ingreso()
        SALIDA()
        Dim j As Integer
        j = DataGridView1.Rows.Count
        For i = 0 To j - 1
            If DataGridView1.Rows(i).Cells(0).Value = True Then
                abrir()
                Dim cmd15 As New SqlCommand("update rama set igsa='1' where partida =@partida  and codigo =@programa", conx)

                cmd15.Parameters.AddWithValue("@partida", Trim(DataGridView1.Rows(i).Cells(8).Value))
                cmd15.Parameters.AddWithValue("@programa", Trim(TextBox1.Text))
                cmd15.ExecuteNonQuery()
            End If

        Next
        TextBox1.Text = ""
        DataGridView1.DataSource = ""
        CheckBox1.Checked = False
    End Sub
    Sub SALIDA()
        Dim fun As New vinsertar_nia
        Dim func As New fnia
        Dim func1 As New insertardetallenia
        Dim func2 As New vnia
        Dim con As String
        Dim hn As New vnia
        Dim fg As New fnia
        Dim func12 As New fnsa
        Dim dts As New vnia
        Dim MES, CORRELATIVO As String
        MES = DateTime.Now.ToString("MM")

        dts.gmes = MES
        dts.galmacen = "06"
        dts.gano = Label2.Text
        dts.gccia = Label3.Text
        CORRELATIVO = func12.buscar_nsa(dts) + 1
        Select Case CORRELATIVO.Length

            Case "1" : CORRELATIVO = "00000" & "" & CORRELATIVO
            Case "2" : CORRELATIVO = "0000" & "" & CORRELATIVO
            Case "3" : CORRELATIVO = "000" & "" & CORRELATIVO
            Case "4" : CORRELATIVO = "00" & "" & CORRELATIVO
            Case "5" : CORRELATIVO = "0" & "" & CORRELATIVO
            Case "6" : CORRELATIVO = CORRELATIVO
        End Select

        MessageBox.Show("LA NOTA DE SALIDA ES  :  " + MES + CORRELATIVO, "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Dim respuesta2 As DialogResult

        respuesta2 = MessageBox.Show("DESEA GENERARLO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta2 = Windows.Forms.DialogResult.Yes) Then
            con = MES & CORRELATIVO
            hn.gcomprobante = con
            hn.galmacen = "06"
            hn.gncom = 2
            hn.gano = Label2.Text
            hn.gccia = Label3.Text
            fg.eliminar_nia(hn)
            Dim ns1 As New vnsacabece
            Dim ns2 As New fnsa

            ns1.gccia = Label3.Text
            ns1.gncom = con
            ns1.gfecha = Format(Now, "dd/MM/yyyy")

            ns1.gorige_sali = "0002"

            ns1.gglosa = "SALIDA A RAMA"

            ns1.gdoc = "002"


            ns1.gguia = TextBox1.Text
            ns1.gsfactu = ""
            ns1.gnfactu = ""
            ns1.gruc = "20508740361"
            ns1.galmacen = "06"
            ns1.gusuario = "HFARJE"
            ns1.gano = Label2.Text

            ns1.gtipointexte = "INT"


            ns1.gtdocento = "002"

            ns1.gfase = ""




            If ns2.insertar_cabecera_nsa(ns1) Then
                Dim i, num2 As Integer
                num2 = DataGridView1.Rows.Count
                For i = 0 To num2 - 1
                    If DataGridView1.Rows(i).Cells(0).Value = True Then
                        Dim nu As String
                        nu = DataGridView1.Rows(i).Cells(9).Value
                        Select Case nu.Length
                            Case "1" : nu = "00" & "" & nu
                            Case "2" : nu = "0" & "" & nu
                            Case "3" : nu = nu
                        End Select
                        Dim ns3 As New vnsadetalle
                        Dim ns4 As New fnsa
                        ns3.gccia = Label3.Text
                        ns3.gncom = MES & CORRELATIVO
                        ns3.gitem = nu
                        ns3.glinea = DataGridView1.Rows(i).Cells(2).Value
                        ns3.gcantidad = DataGridView1.Rows(i).Cells(6).Value
                        ns3.gund = DataGridView1.Rows(i).Cells(7).Value
                        If DataGridView1.Rows(i).Cells(1).Value = "" Then
                            ns3.gop = ""
                        Else
                            ns3.gop = DataGridView1.Rows(i).Cells(1).Value
                        End If
                        ns3.gunidad_medidad = DataGridView1.Rows(i).Cells(5).Value
                        ns3.gpartida = DataGridView1.Rows(i).Cells(8).Value
                        ns3.grollo1 = DataGridView1.Rows(i).Cells(4).Value
                        ns3.galmacen = "06"
                        ns3.gordtejeduria = ""
                        ns3.gano = Label2.Text
                        If Convert.ToString(DataGridView1.Rows(i).Cells(8).Value) = "" Then
                            ns3.glote = ""
                        Else
                            ns3.glote = DataGridView1.Rows(i).Cells(8).Value
                        End If

                        ns3.gcantenvio = "0.00"



                        ns4.insertar_detalle_nsa(ns3)
                    End If


                Next
                MessageBox.Show("Datos registrados correctamente ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Dim respuesta As DialogResult

                respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim jp As New Reporte_Nia_Nsa


                    jp.TextBox1.Text = "06"
                    jp.TextBox2.Text = 2
                    jp.TextBox3.Text = con
                    jp.TextBox4.Text = Label2.Text
                    jp.TextBox5.Text = Label3.Text

                    jp.Show()
                End If
            End If
        End If

    End Sub

    Sub ingreso()
        Dim fun As New vinsertar_nia
        Dim func As New fnia
        Dim func1 As New insertardetallenia
        Dim func2 As New fnia

        Dim hn As New vnia
        Dim fg As New fnia
        Dim func12 As New fnia
        Dim dts As New vnia
        Dim mes, correlativo As String
        mes = DateTime.Now.ToString("MM")
        dts.gccia = Label3.Text
        dts.gmes = mes
        dts.galmacen = "06"
        dts.gano = Label2.Text

        correlativo = func12.buscar_nia(dts) + 1
        Select Case correlativo.Length

            Case "1" : correlativo = "00000" & "" & correlativo
            Case "2" : correlativo = "0000" & "" & correlativo
            Case "3" : correlativo = "000" & "" & correlativo
            Case "4" : correlativo = "00" & "" & correlativo
            Case "5" : correlativo = "0" & "" & correlativo
            Case "6" : correlativo = correlativo
        End Select

        MessageBox.Show("LA NOTA DE INGRESO ES  " + mes + correlativo, "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Dim respuesta2 As DialogResult

        respuesta2 = MessageBox.Show("DESEA GENERARLO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta2 = Windows.Forms.DialogResult.Yes) Then
            fun.gccia = Label3.Text
            fun.gncom = mes + correlativo
            fun.gglosa = "PARTE DIARIO TEJEDURIA"
            fun.gfecha = Format(Now, "dd/MM/yyyy")
            fun.gguia = Trim(TextBox1.Text)
            fun.gano = Label2.Text

            fun.gdoc = "002"

            fun.galmacen = "06"
            fun.gempresa = "20508740361"

            fun.gint = "INT"


            fun.gorigen = "0001"

            fun.gtidoc = "002"

            fun.gusuario = "HFARJE"
            fun.gfase = ""
            If func.insertar_cabecera_nia(fun) Then

                Dim i, num2 As Integer

                num2 = DataGridView1.Rows.Count
                For i = 0 To num2 - 1
                    If DataGridView1.Rows(i).Cells(0).Value = True Then
                        Dim nu As String
                        Dim jj As New fingresopac
                        Dim aa As New vpackingtela
                        func1.gccia = Label3.Text
                        func1.gncom = mes + correlativo

                        nu = DataGridView1.Rows(i).Cells(9).Value

                        Select Case nu.Length
                            Case "1" : nu = "00" & "" & nu
                            Case "2" : nu = "0" & "" & nu
                            Case "3" : nu = nu
                        End Select
                        func1.gitem = nu

                        func1.glinea = DataGridView1.Rows(i).Cells(2).Value
                        If DataGridView1.Rows(i).Cells(1).Value = "" Then
                            func1.gop = ""
                        Else
                            func1.gop = DataGridView1.Rows(i).Cells(1).Value
                        End If

                        func1.gund = DataGridView1.Rows(i).Cells(7).Value
                        func1.gcantidad = DataGridView1.Rows(i).Cells(4).Value
                        func1.gcantidad1 = DataGridView1.Rows(i).Cells(6).Value
                        If DataGridView1.Rows(i).Cells(8).Value = "" Then
                            func1.gpartida = ""
                        Else
                            func1.gpartida = DataGridView1.Rows(i).Cells(8).Value
                        End If

                        func1.galmacen = "06"
                        func1.gumpresentacion = DataGridView1.Rows(i).Cells(5).Value
                        func1.gotejeduria = ""


                        func1.goc = ""

                        func1.gano = Label2.Text
                        If DataGridView1.Rows(i).Cells(8).Value = "" Then
                            func1.glote = ""
                        Else
                            func1.glote = DataGridView1.Rows(i).Cells(8).Value
                        End If

                        func1.gcenvio = "0.00"

                        func2.insertar_detalle_nia(func1)
                    End If

                Next
                MessageBox.Show("Datos registrados correctamente ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Dim respuesta As DialogResult

                respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE INGRESO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then

                    Reporte_Nia_Nsa.TextBox1.Text = "06"
                    Reporte_Nia_Nsa.TextBox2.Text = 1
                    Reporte_Nia_Nsa.TextBox3.Text = mes + correlativo
                    Reporte_Nia_Nsa.TextBox4.Text = Label2.Text
                    Reporte_Nia_Nsa.TextBox5.Text = Label3.Text
                    Reporte_Nia_Nsa.Show()
                End If



            End If
        End If

    End Sub

    Private Sub Generara_ingreso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class