Imports System.Data.SqlClient
Public Class Otejeduria
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand

    Private Sub Otejeduria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Rsr2, Rsr20, Rsr21, Rsr22 As SqlDataReader
        RadioButton3.Checked = True
        abrir()
        llenar_combo_box2()
        llenar_combo_box3()
        If Label17.Text = 1 Then
            Dim sql102 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAllFichaTextilTelaDetalle '" + Trim(Label14.Text) + "','" + Trim(DataGridView1.Rows(0).Cells(0).Value) + "'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            Dim i5 As Integer
            i5 = 0

            While Rsr2.Read()
                DataGridView2.Rows.Add()
                DataGridView2.Rows(i5).Cells(0).Value = Rsr2(2)
                DataGridView2.Rows(i5).Cells(1).Value = Rsr2(13)
                DataGridView2.Rows(i5).Cells(2).Value = Rsr2(12)
                DataGridView2.Rows(i5).Cells(7).Value = i5 + 1
                i5 = i5 + 1
                'DataGridView2.Rows(i5).Cells(6).Value = (Rsr2(12) * DataGridView1.Rows(0).Cells(10).Value) / 100
            End While
            Rsr2.Close()

        Else
            If Label17.Text = 2 Then
                Dim PO As String
                PO = TextBox1.Text + TextBox2.Text

                Dim sql1020 As String = "SELECT A.ncom_4,A.fcom_4,A.gloa_4,A.prvtej_4,A.acabado_4,A.presen_4,A.psxrolst_4,A.nrevrol_4,A.rpmrol_4,A.prvtin_4,A.partidA_4,A.color_4,       
		       ISNULL(C.CantReq_4a,0) AS CantReq_4a,E.cnomb_3c	    
		       FROM custom_vianny.dbo.matg040f A  LEFT JOIN custom_vianny.dbo.Matp040f C
		       					ON A.ccia_4= C.CCia_4a AND ncom_4 = C.NCom_4a 
		    			   		LEFT OUTER JOIN custom_vianny.dbo.Cag1700 D
		    			   		ON C.CCia_4a = D.CCia AND C.Tela_4a  =D.Linea_17	
								LEFT  JOIN custom_vianny.dbo.Qarc0300 E
								ON A.color_4 = E.ccolor_3c  AND A.ccia_4 =E.ccia_3c
		         Where A.ccia_4='01' AND A.ncom_4='" + PO + "'"
                Dim cmd1020 As New SqlCommand(sql1020, conx)
                Rsr20 = cmd1020.ExecuteReader()
                Rsr20.Read()
                DateTimePicker1.Value = Rsr20(1)
                TextBox10.Text = Rsr20(2)
                TextBox8.Text = Rsr20(3)
                ComboBox1.Text = Rsr20(4).ToString
                ComboBox2.Text = Rsr20(5).ToString
                TextBox13.Text = Rsr20(6)
                TextBox14.Text = Rsr20(7)
                TextBox11.Text = Rsr20(8)
                TextBox3.Text = Rsr20(9)
                TextBox5.Text = Rsr20(10)
                TextBox6.Text = Rsr20(11)
                TextBox7.Text = Rsr20(13)

                DataGridView1.Rows(0).Cells(10).Value = Rsr20(12)
                Rsr20.Close()

                Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAllFichaTextilTelaDetalle '" + Trim(Label14.Text) + "','" + Trim(DataGridView1.Rows(0).Cells(0).Value) + "'"
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr21 = cmd1021.ExecuteReader()
                Dim i51 As Integer
                i51 = 0

                While Rsr21.Read()
                    DataGridView2.Rows.Add()
                    DataGridView2.Rows(i51).Cells(0).Value = Rsr21(2)
                    DataGridView2.Rows(i51).Cells(1).Value = Rsr21(13)
                    DataGridView2.Rows(i51).Cells(2).Value = Rsr21(12)
                    i51 = i51 + 1
                End While
                Rsr21.Close()
                Dim O As Integer
                O = DataGridView2.Rows.Count
                For i = 0 To O - 1
                    Dim sql1022 As String = " SELECT prvhil_41a,lote_41a,cant_41a,itm_41a FROM custom_vianny.dbo.matp041f A WHERE  A.ncom_41a='" + PO + "' AND hilo_41a ='" + Trim(DataGridView2.Rows(i).Cells(0).Value) + "'"
                    Dim cmd1022 As New SqlCommand(sql1022, conx)
                    Rsr22 = cmd1022.ExecuteReader()
                    Rsr22.Read()
                    DataGridView2.Rows(i).Cells(3).Value = Rsr22(0)
                    DataGridView2.Rows(i).Cells(5).Value = Rsr22(1)
                    DataGridView2.Rows(i).Cells(6).Value = Rsr22(2)
                    DataGridView2.Rows(i).Cells(7).Value = Rsr22(3)
                    Rsr22.Close()
                Next
                Button6.Enabled = False
                Button2.Enabled = False
                Button1.Enabled = False
                DataGridView1.Enabled = False
                DataGridView2.Enabled = False
                TextBox3.Enabled = False
                TextBox8.Enabled = False
                TextBox10.Enabled = False
                TextBox11.Enabled = False
                TextBox13.Enabled = False
                TextBox14.Enabled = False
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
            End If
        End If

    End Sub
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("EXEC custom_vianny.dbo.CAESOFT_GetAllPresentacionRollotejeduria", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "DESCRIPCION"
            ComboBox1.ValueMember = "CODIGO"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Sub llenar_combo_box3()
        Try

            conn = New SqlDataAdapter("EXEC custom_vianny.dbo.CAESOFT_GetAllAcabadoRollotejeduria", conx)
            conn.Fill(TY3)
            ComboBox2.DataSource = TY3
            ComboBox2.DisplayMember = "DESCRIPCIOn"
            ComboBox2.ValueMember = "CODIGO"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = "39"
            Clientes.Show()
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If e.ColumnIndex = 4 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            Lotes.Label3.Text = DataGridView2.Rows(num1).Cells(0).Value
            Lotes.Label4.Text = Label13.Text
            Lotes.Label5.Text = num1
            Lotes.Show()

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Trim(DataGridView1.Rows(0).Cells(0).Value).Length <> 0 Then
            Dim Rsr19 As SqlDataReader
            Dim sql101 As String = "EXEC custom_vianny.dbo.spu_jaladatostecnicosfichaparaOT '" + Trim(Label14.Text) + "','" + Trim(DataGridView1.Rows(0).Cells(0).Value) + "'"
            Dim cmd101 As New SqlCommand(sql101, conx)
            Rsr19 = cmd101.ExecuteReader()
            Rsr19.Read()
            DataGridView1.Rows(0).Cells(2).Value = Rsr19(1)
            DataGridView1.Rows(0).Cells(3).Value = Rsr19(2)
            DataGridView1.Rows(0).Cells(4).Value = Rsr19(3)
            DataGridView1.Rows(0).Cells(5).Value = Rsr19(4)
            DataGridView1.Rows(0).Cells(6).Value = Rsr19(5)
            DataGridView1.Rows(0).Cells(7).Value = Rsr19(6)
            Rsr19.Close()
            Dim p As Integer
            p = DataGridView2.Rows.Count
            For i = 0 To p - 1

                DataGridView2.Rows(i).Cells(6).Value = (DataGridView2.Rows(i).Cells(2).Value * DataGridView1.Rows(0).Cells(10).Value) / 100
            Next

        Else
            MsgBox("NO EXISTE CODIO DE TELA PARA BUSCAR SU FICHA TECNICA")
        End If

    End Sub

    Private Sub Otejeduria_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        EMISION_PARTIDAS.Button6.Enabled = True
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim h, h1, Ph, Lt, Dt As Integer
        h = DataGridView2.Rows.Count
        h1 = DataGridView1.Rows.Count
        Ph = 0
        Lt = 0
        Dt = 0
        For i = 0 To h - 1
            If Trim(DataGridView2.Rows(i).Cells(3).Value).Length = 0 Then
                Ph = Ph + 1
            End If
            If Trim(DataGridView2.Rows(i).Cells(5).Value).Length = 0 Then
                Lt = Lt + 1
            End If
        Next
        For i1 = 0 To h1 - 1
            If Trim(DataGridView2.Rows(i1).Cells(2).Value).Length = 0 Then
                Dt = Dt + 1
            End If

        Next

        If h = 0 Then
            MessageBox.Show("NO HAY INFORMACION DEL HILO REQUERIDO", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            If Ph > 0 Then
                MessageBox.Show("FALTA INFORMACION DE UN PROVEEDOR EN ALGUNA CELDA", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                If Lt > 0 Then
                    MessageBox.Show("FALTA INFORMACION DEL LOTE EN ALGUNA CELDA", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Else
                    If Dt > 0 Then
                        MessageBox.Show("FALTA INFORMACION DE LOS DATOS TECNICOS", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Else
                        If Trim(TextBox13.Text).Length = 0 Then
                            MessageBox.Show("ES OBLIGATORIO INGRESOS EL PESO POR ROLLO", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Else

                            abrir()
                            Dim agregar As String = "DELETE FROM custom_vianny.dbo.Qanp300 WHERE ccia_3p  = '" + Trim(Label14.Text) + "' AND regis_3p= '" + Trim(TextBox5.Text) + "'"
                            Dim agregar1 As String = "delete from custom_vianny.dbo.Ranp300 where regis_3p ='" + Trim(TextBox5.Text) + "'"
                            ELIMINAR(agregar)
                            ELIMINAR(agregar1)
                            Dim cmd15 As New SqlCommand("INSERT INTO custom_vianny.dbo.Ranp300 (ccia_3p , regis_3p , Op_3p, linea_3p , KgArmados_3p)VALUES ( @ccia_3p , @regis_3p, @Op_3p, @linea_3p,@KgArmados_3p)", conx)
                            cmd15.Parameters.AddWithValue("@ccia_3p", Trim(Label14.Text))
                            cmd15.Parameters.AddWithValue("@regis_3p", Trim(TextBox5.Text))
                            cmd15.Parameters.AddWithValue("@Op_3p", Trim(TextBox12.Text))
                            cmd15.Parameters.AddWithValue("@linea_3p", Trim(DataGridView1.Rows(0).Cells(0).Value))
                            cmd15.Parameters.AddWithValue("@KgArmados_3p", Trim(DataGridView1.Rows(0).Cells(10).Value))
                            cmd15.ExecuteNonQuery()
                            guardar()


                            Dim respuesta As DialogResult
                            respuesta = MessageBox.Show("DESEAS ACTUALIZAR INFORMACION DE LOS ROLLOS?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                                Dim cmd16 As New SqlCommand("update custom_vianny.dbo.marg0001 set maquina_3r =@maquina, galga_3r=@galga, ancho_3r=@ancho, diam_3r=@diametro,ccolor_3r=@color where partida_3r =@partida and ccia_3r ='01'", conx)
                                cmd16.Parameters.AddWithValue("@maquina", Trim(DataGridView1.Rows(0).Cells(1).Value))
                                cmd16.Parameters.AddWithValue("@galga", Trim(DataGridView1.Rows(0).Cells(3).Value))
                                cmd16.Parameters.AddWithValue("@ancho", Trim(DataGridView1.Rows(0).Cells(4).Value))
                                cmd16.Parameters.AddWithValue("@diametro", Trim(DataGridView1.Rows(0).Cells(2).Value))
                                cmd16.Parameters.AddWithValue("@color", Trim(TextBox6.Text))
                                cmd16.Parameters.AddWithValue("@partida", Trim(TextBox5.Text))

                                cmd16.ExecuteNonQuery()

                                MsgBox("SE ACTUALIZO LA INFORMACION")
                            End If
                        End If
                    End If
                End If
            End If
        End If


    End Sub
    Sub guardar()
        'Try
        Dim kp As New fprogramacion
            Dim jk As New vmatg040f
            Dim jk1 As New vmatp040f
            Dim jk3 As New vqan0300
            Dim jk4 As New vqanp300
            Dim jk2 As New vmatp041f
        jk.gccia_4 = Trim(Label14.Text)
        jk.gncom4 = TextBox1.Text + TextBox2.Text
        jk.gfecha = DateTimePicker1.Value
        jk.gobservacion = Trim(TextBox10.Text)
        jk.gusuario = ""
        jk.gpartida_4 = Trim(TextBox5.Text)
        jk.gcolor_4 = Trim(TextBox6.Text)
        jk.gprvtin_4 = Trim(TextBox8.Text)
        jk.gprvtej_4 = Trim(TextBox3.Text)
        jk.gacabado_4 = ComboBox2.Text
        jk.gpresen_4 = ComboBox1.Text
        If TextBox13.Text = "" Then
                jk.gpsxrolst_4 = 0
            Else
                jk.gpsxrolst_4 = TextBox13.Text
            End If
            If TextBox14.Text = "" Then
                jk.gnrevrol_4 = 0
            Else
                jk.gnrevrol_4 = TextBox14.Text
            End If
            If TextBox11.Text = "" Then
                jk.grpmrol_4 = 0
            Else
                jk.grpmrol_4 = TextBox11.Text
            End If

            kp.insertar_matg040f(jk)

        jk1.gccia_4a = Trim(Label14.Text)
            jk1.gncom_4a = TextBox1.Text + TextBox2.Text
            jk1.gTELA_4A = DataGridView1.Rows(0).Cells(0).Value

        If Trim(DataGridView1.Rows(0).Cells(2).Value) = "" Then
            jk1.gDIAMET_4A = "0.00"
        Else
            jk1.gDIAMET_4A = DataGridView1.Rows(0).Cells(2).Value
            End If

        If Trim(DataGridView1.Rows(0).Cells(3).Value) = "" Then
            jk1.gGALGA_4A = "0.00"
        Else
            jk1.gGALGA_4A = DataGridView1.Rows(0).Cells(3).Value
            End If
        If Trim(DataGridView1.Rows(0).Cells(10).Value) = "" Then
            jk1.gCANTREQ_4A = "0.00"
        Else
            jk1.gCANTREQ_4A = DataGridView1.Rows(0).Cells(10).Value
            End If

            jk1.gMAQUINA_4A = DataGridView1.Rows(0).Cells(1).Value
            jk1.gANCHO_4A = DataGridView1.Rows(0).Cells(4).Value

        If Trim(DataGridView1.Rows(0).Cells(5).Value) = "" Then
            jk1.gLM1_41A = "0.00"
        Else
            jk1.gLM1_41A = DataGridView1.Rows(0).Cells(5).Value
            End If
        If Trim(DataGridView1.Rows(0).Cells(6).Value) = "" Then
            jk1.gLM2_41A = "0.00"
        Else
            jk1.gLM2_41A = DataGridView1.Rows(0).Cells(6).Value
            End If
        If Trim(DataGridView1.Rows(0).Cells(7).Value) = "" Then
            jk1.gLM3_41A = "0.00"
        Else
            jk1.gLM3_41A = DataGridView1.Rows(0).Cells(7).Value
            End If
        If Trim(DataGridView1.Rows(0).Cells(8).Value) = "" Then
            jk1.gLM4_41A = "0.00"
        Else
            jk1.gLM4_41A = DataGridView1.Rows(0).Cells(8).Value
            End If
        If Trim(DataGridView1.Rows(0).Cells(9).Value) = "" Then
            jk1.gLM5_41A = "0.00"
        Else
            jk1.gLM5_41A = DataGridView1.Rows(0).Cells(9).Value
            End If

        kp.insertar_matp040f(jk1)

            Dim P, M As Integer
            M = 1
            P = DataGridView2.Rows.Count
            For I1 = 0 To P - 1

                jk2.gccia_41A = Trim(Label14.Text)
                jk2.gncom_41a = TextBox1.Text + TextBox2.Text
                Select Case M.ToString.Length
                    Case "1" : M = "0" & "" & M
                    Case "2" : M = "0" & "" & M
                    Case "3" : M = M
                End Select
            jk2.gitm_41a = DataGridView2.Rows(I1).Cells(7).Value
            jk2.ghilo_41a = DataGridView2.Rows(I1).Cells(0).Value

            If Trim(DataGridView1.Rows(0).Cells(6).Value) = "" Then
                jk2.gcant_41a = 0
            Else
                jk2.gcant_41a = DataGridView2.Rows(I1).Cells(6).Value
                End If
                jk2.glote_41a = DataGridView2.Rows(I1).Cells(5).Value
            jk2.gprvhil_41a = DataGridView2.Rows(I1).Cells(3).Value

            kp.insertar_matp041f(jk2)
                M = M + 1
            Next


            jk3.gccia_3n = Trim(Label14.Text)
            jk3.gregis_3n = Trim(TextBox5.Text)
            jk3.gncom_3n = TextBox1.Text
            jk3.gfecha = DateTimePicker1.Value
            jk3.gorigen_3n = "INT"
            jk3.gFich_3n = Trim(TextBox3.Text)
            jk3.gcolor_3n = Trim(TextBox6.Text)
            jk3.gfase_3n = "0201"
        jk3.gestado_3n = "01"
        jk3.gGlosa1_3n = Trim(TextBox10.Text)
        kp.insertar_qan0300(jk3)

            jk4.gccia_3p = Trim(Label14.Text)
            jk4.gregis_3p = Trim(TextBox5.Text)
            jk4.gOp_3p = Trim(TextBox12.Text)
            jk4.glinea_3p = DataGridView1.Rows(0).Cells(0).Value
            jk4.ggalga_3p = DataGridView1.Rows(0).Cells(3).Value
            jk4.gAncho_3p = DataGridView1.Rows(0).Cells(4).Value
            jk4.gkgreq_3p = DataGridView1.Rows(0).Cells(10).Value
        jk4.gobserv_3p = ""

        kp.insertar_qanp300(jk4)
            MsgBox("LA INFORMACION SE REGISTRO CORRECTAMENTE")
        'Catch ex As Exception
        '    MsgBox("NO SE REGISTRO LA INFORMACION")
        'End Try

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Crear_rollos.TextBox1.Text = TextBox1.Text
        Crear_rollos.TextBox2.Text = TextBox2.Text
        Crear_rollos.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button2.Enabled = True
        Button1.Enabled = True
        DataGridView1.Enabled = True
        DataGridView2.Enabled = True
        TextBox3.Enabled = True
        TextBox8.Enabled = True
        TextBox10.Enabled = True
        TextBox11.Enabled = True
        TextBox13.Enabled = True
        TextBox14.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        DataGridView1.BeginEdit(True)
        DataGridView2.BeginEdit(True)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ROLLOS_ADICIONALES.TextBox1.Text = TextBox1.Text
        ROLLOS_ADICIONALES.TextBox2.Text = TextBox2.Text
        ROLLOS_ADICIONALES.Label11.Text = DataGridView1.Rows(0).Cells(1).Value
        ROLLOS_ADICIONALES.Label13.Text = TextBox3.Text
        ROLLOS_ADICIONALES.Label14.Text = TextBox8.Text
        ROLLOS_ADICIONALES.Show()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        lista_rollos.Label3.Text = TextBox5.Text
        lista_rollos.Label4.Text = TextBox1.Text & TextBox2.Text
        lista_rollos.Show()
        'TEJIDO.TextBox1.Text = TextBox5.Text
        'TEJIDO.Show()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Rpt_Tejeduria.TextBox1.Text = TextBox1.Text & TextBox2.Text
        Rpt_Tejeduria.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = "40"
            Clientes.Show()
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.ColumnIndex = 1 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            Det_Rollo.Label1.Text = "TEJEDORA"
            Det_Rollo.Label2.Text = 45
            Det_Rollo.Label3.Text = num1
            Det_Rollo.ShowDialog()

        End If
    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        If e.ColumnIndex = 3 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            Clientes.TextBox3.Text = "339"
            Clientes.Label3.Text = num1
            Clientes.ShowDialog()

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        abrir()
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("SE ELIMINAR TODA LA INFORMACIO DE LA ORDEN DE TEJIDO, PARTIDA Y ROLLOS AGREGADOS ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Dim agregar As String = "DELETE FROM custom_vianny.dbo.Qanp300 WHERE ccia_3p  = '" + Trim(Label14.Text) + "' AND regis_3p= '" + Trim(TextBox5.Text) + "'"
            Dim agregar1 As String = "delete from custom_vianny.dbo.Ranp300 where regis_3p ='" + Trim(TextBox5.Text) + "'"

            Dim agregar2 As String = "delete From custom_vianny.dbo.matp040f Where ncom_4a ='" + TextBox1.Text & TextBox2.Text + "'   and ccia_4a ='01'"
            Dim agregar3 As String = "delete From custom_vianny.dbo.matg040f Where ncom_4 ='" + TextBox1.Text & TextBox2.Text + "'  and ccia_4 ='01'"
            Dim agregar4 As String = "delete From custom_vianny.dbo.matp041f Where ncom_41a ='" + TextBox1.Text & TextBox2.Text + "' and ccia_41a ='01'"
            Dim agregar5 As String = "delete From custom_vianny.dbo.qan0300 Where regis_3n ='" + Trim(TextBox5.Text) + "'  and ccia_3n ='01'"
            Dim agregar6 As String = "delete From custom_vianny.dbo.marg0001 Where partida_3r ='" + Trim(TextBox5.Text) + "' AND ccia_3r ='01' "

            ELIMINAR(agregar)
            ELIMINAR(agregar1)
            ELIMINAR(agregar2)
            ELIMINAR(agregar3)
            ELIMINAR(agregar4)
            ELIMINAR(agregar5)
            ELIMINAR(agregar6)
            MsgBox("SE ELIMINO LA INFORMACION CORRECTAMENTE")
            Me.Close()
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Otejeduria_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        EMISION_PARTIDAS.DataGridView1.BeginEdit(True)
        abrir()
        EMISION_PARTIDAS.DataGridView3.DataSource = ""
        EMISION_PARTIDAS.DataGridView2.DataSource = ""
        Dim dg1 As New DataTable
        Dim cmd As New SqlDataAdapter("select regis_3p as PARTIDA,QG.fcom_3n AS FECHA,c.nomb_17 AS CODIGO, QP.kgreq_3p  from custom_vianny.dbo.Qanp300 QP 
					   INNER JOIN custom_vianny.dbo.qan0300 QG ON QP.ccia_3p = QP.ccia_3p AND QP.regis_3p =QG.regis_3n 
					   LEFT JOIN custom_vianny.dbo.cag1700 c on qp.ccia_3p = c.ccia and qp.linea_3p = c.linea_17
					   where Op_3p ='" + Trim(TextBox12.Text) + "' and qg.flag_3n =1 order by regis_3p", conx)

        cmd.Fill(dg1)
        If dg1.Rows.Count <> 0 Then
            EMISION_PARTIDAS.DataGridView3.DataSource = dg1
            EMISION_PARTIDAS.DataGridView3.Columns(2).Width = 550
        Else
            EMISION_PARTIDAS.DataGridView3.DataSource = ""
        End If

        Dim dg2 As New DataTable
        Dim cmd2 As New SqlDataAdapter("Select ncom_4a as OT,G.partidA_4 AS PARTIDA,G.fcom_4 AS FECHA,P.cantreq_4a AS KGS_PROGRAMADO,P.maquina_4a AS MAQ,P.diamet_4a AS D,P.galga_4a G,P.ancho_4a AS A,P.lm1_41a AS LM1, P.lm2_41a AS LM2,P.lm3_41a AS LM2, P.tela_4a  
			 from  custom_vianny.dbo.matp040f P INNER JOIN custom_vianny.dbo.matg040f G ON P.ncom_4a = G.ncom_4 AND G.ccia_4 = P.ccia_4a 
			 where SUBSTRING(ncom_4a,1,10) ='" + Trim(TextBox1.Text) + "' AND G.flag_4=1 order by partidA_4", conx)

        cmd2.Fill(dg2)
        If dg2.Rows.Count <> 0 Then
            EMISION_PARTIDAS.DataGridView2.DataSource = dg2
            'DataGridView3.Columns(2).Width = 350
            EMISION_PARTIDAS.DataGridView2.Columns(11).Visible = False
        Else
            EMISION_PARTIDAS.DataGridView2.DataSource = ""
        End If
        EMISION_PARTIDAS.Label6.Text = Trim(TextBox12.Text)
    End Sub
End Class