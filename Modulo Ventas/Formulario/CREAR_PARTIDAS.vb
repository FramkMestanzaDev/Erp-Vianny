Imports System.Data.SqlClient
Public Class CREAR_PARTIDAS
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3, TY4 As New DataTable
    Public comando As SqlCommand
    Dim Rsr2, Rsr25, Rsr26, Rsr27 As SqlDataReader
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
    Sub ULTIMO()
        abrir()
        Dim sql102 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAllUltimaPartida '" + Label12.Text + "','" + Label11.Text + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        Dim i5 As Integer
        i5 = 0
        Rsr2.Read()
        TextBox1.Text = Rsr2(0)
        Rsr2.Close()
        ComboBox2.SelectedIndex = 0
        ComboBox1.SelectedIndex = 0


    End Sub
    Private Sub CREAR_PARTIDAS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Label14.Text = 2 Then
            abrir()
            llenar_combo_box3()
            ULTIMO()
        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Clear()
            abrir()
            Dim sql1022 As String = "select COUNT(regis_3n) from  custom_vianny.DBO.qan0300 where regis_3n ='" + Trim(TextBox1.Text) + "'  and ccia_3n ='" + Label12.Text + "'"
            Dim cmd1022 As New SqlCommand(sql1022, conx)
            Rsr25 = cmd1022.ExecuteReader()
            Dim i5 As Integer
            i5 = 0
            Rsr25.Read()
            If Rsr25(0) > 0 Then
                abrir()
                Dim sql1023 As String = "SELECT A.fcom_3n,A.origen_3n,a.color_3n,a.flag_3n,a.fich_3n,a.tipotela_3n,a.estado_3n,
		       		  ISNULL(C.Nomb_10,'') AS Nomb_10,
		       		  ISNULL(D.CNomb_3c,'') AS CNomb_3c,
		       		  ISNULL(E.Nomb_4,'') AS Nomb_4,ncom_3n,glosa1_3n,fase_3n
		       		  FROM custom_vianny.DBO.qan0300 A  LEFT JOIN custom_vianny.DBO.Cag1000 C
		       		  				ON A.CCia_3n = C.CCia AND A.Fich_3n = C.Fich_10
		       		  				  LEFT JOIN custom_vianny.DBO.Qarc0300 D
		       		  				ON A.CCia_3n = D.CCia_3c AND A.Color_3n = D.CColor_3c
		       		  				  LEFT JOIN custom_vianny.DBO.Qag0400 E
		       		  				ON A.CCia_3n = E.CCia AND A.Fase_3n = E.Fase_4
		       		   Where CCia_3n = '" + Label12.Text + "' AND Regis_3n = '" + Trim(TextBox1.Text) + "'  
		       		   Order By A.ccia_3N , A.regis_3n"
                Dim cmd1023 As New SqlCommand(sql1023, conx)
                Rsr26 = cmd1023.ExecuteReader()

                Rsr26.Read()
                TextBox2.Text = Rsr26(10)
                TextBox3.Text = Rsr26(2)
                TextBox4.Text = Rsr26(8)
                TextBox5.Text = Rsr26(4)
                TextBox6.Text = Rsr26(7)
                TextBox7.Text = Rsr26(11)
                If Rsr26(5) = "TC" Then
                    ComboBox1.SelectedIndex = 0
                Else
                    ComboBox1.SelectedIndex = 1
                End If
                If Rsr26(6) = "01" Then
                    ComboBox2.SelectedIndex = 0
                Else
                    ComboBox2.SelectedIndex = 1
                End If

                If Rsr26(1) = "INT" Then
                    ComboBox3.SelectedIndex = 0
                Else
                    ComboBox3.SelectedIndex = 1
                End If
                DateTimePicker1.Value = Rsr26(0)
                Label13.Text = Rsr26(12)
                Rsr26.Close()
                Dim sql102 As String = "SELECT A.ccia_3p,regis_3p,linea_17,a.galga_3p,a.kgreq_3p,a.ancho_3p,a.Op_3p ,
		       		  ISNULL(B.Nomb_17,'') AS Nomb_17
		       		  FROM custom_vianny.DBO.qanp300 A  LEFT JOIN custom_vianny.DBO.Cag1700 B
		       					ON '01' = B.CCia AND A.Linea_3p = B.Linea_17
		       		   Where CCia_3p = '" + Label12.Text + "' AND Regis_3p = '" + Trim(TextBox1.Text) + "' 
		       		   Order By A.regis_3p"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr27 = cmd102.ExecuteReader()
                Rsr27.Read()
                DataGridView1.Rows.Add()
                DataGridView1.Rows(0).Cells(0).Value = Rsr27(6)
                DataGridView1.Rows(0).Cells(1).Value = Rsr27(2)
                DataGridView1.Rows(0).Cells(2).Value = Rsr27(7)
                DataGridView1.Rows(0).Cells(3).Value = Rsr27(3)
                DataGridView1.Rows(0).Cells(4).Value = Rsr27(5)
                DataGridView1.Rows(0).Cells(5).Value = Rsr27(4)
                Rsr27.Close()

                Rsr25.Close()
                abrir()
                llenar_combo_box4(Label13.Text.ToString)
                Button1.Enabled = False
            Else
                ComboBox3.SelectedIndex = 0
                TextBox5.Text = "20508740361"
                TextBox6.Text = "CONSORCIO TEXTIL VIANNYSA S.A.C."
                TextBox5.Enabled = True
                TextBox2.Enabled = True
                TextBox7.ReadOnly = False
                DataGridView1.Enabled = True
                ComboBox3.Enabled = True
                ComboBox4.Enabled = True
                TextBox2.Select()
            End If






        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        REQUERIMIENTO.Label1.Text = TextBox2.Text
        REQUERIMIENTO.Show()
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            Det_Rollo.Label1.Text = "PO"
            Det_Rollo.Label2.Text = 44
            Det_Rollo.Show()
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.Rows.Clear()
        TextBox7.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        TextBox3.Text = ""
        TextBox2.Text = ""
        TextBox7.ReadOnly = True
        TextBox1.Select()
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
    Sub llenar_combo_box3()
        Try

            conn = New SqlDataAdapter("SELECT fase_4,nomb_4 FROM custom_vianny.dbo.Qag0400 A  Where A.ccia = '" + Label12.Text + "' AND substring(fase_4,1,2)='02'", conx)
            conn.Fill(TY3)
            ComboBox4.DataSource = TY3
            ComboBox4.DisplayMember = "nomb_4"
            ComboBox4.ValueMember = "fase_4"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox5.Enabled = True
        TextBox2.Enabled = True
        TextBox7.ReadOnly = False
        DataGridView1.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        Button1.Enabled = True
        Button4.Enabled = True
        TextBox2.Select()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Rpte_Partida.TextBox1.Text = TextBox1.Text
        Rpte_Partida.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim agregar1 As String = "DELETE FROM custom_vianny.dbo.Qanp300 WHERE ccia_3p  = '" + Label12.Text + "' AND regis_3p= '" + Trim(TextBox1.Text) + "'"
        Dim agregar2 As String = "delete From custom_vianny.dbo.qan0300 Where regis_3n ='" + Trim(TextBox1.Text) + "' and ccia_3n ='" + Label12.Text + "'"
        ELIMINAR(agregar1)
        ELIMINAR(agregar2)
        MsgBox("SE ELIMINO LA INFORMACION SOLICITADA")
        ULTIMO()
        DataGridView1.Rows.Clear()
        TextBox7.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        TextBox3.Text = ""
        TextBox2.Text = ""
        TextBox7.ReadOnly = True
        TextBox1.Select()

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Sub llenar_combo_box4(AB As String)
        Try

            conn = New SqlDataAdapter("SELECT fase_4,nomb_4 FROM custom_vianny.dbo.Qag0400 A  Where A.ccia = '" + Label12.Text + "' AND substring(fase_4,1,2)='02' AND fase_4='" + Trim(AB) + "' ", conx)
            conn.Fill(TY4)
            ComboBox4.DataSource = TY4
            ComboBox4.DisplayMember = "nomb_4"
            ComboBox4.ValueMember = "fase_4"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim jk3 As New vqan0300
        Dim jk4 As New vqanp300
        Dim kp As New fprogramacion
        Dim agregar1 As String = "DELETE FROM custom_vianny.dbo.Qanp300 WHERE ccia_3p  = '" + Label12.Text + "' AND regis_3p= '" + Trim(TextBox1.Text) + "'"
        Dim agregar2 As String = "delete From custom_vianny.dbo.qan0300 Where regis_3n ='" + Trim(TextBox1.Text) + "' and ccia_3n ='" + Label12.Text + "'"
        ELIMINAR(agregar1)
        ELIMINAR(agregar2)
        jk3.gccia_3n = Trim(Label12.Text)
        jk3.gregis_3n = Trim(TextBox1.Text)
        jk3.gncom_3n = TextBox2.Text
        jk3.gfecha = DateTimePicker1.Value
        If ComboBox3.Text = "INTERNA" Then
            jk3.gorigen_3n = "INT"
        Else
            jk3.gorigen_3n = "EXT"
        End If

        jk3.gFich_3n = Trim(TextBox5.Text)
        jk3.gcolor_3n = Trim(TextBox3.Text)
        jk3.gfase_3n = ComboBox4.SelectedValue.ToString
        If ComboBox3.Text = "PRODUCCION" Then
            jk3.gestado_3n = "01"
        Else
            jk3.gestado_3n = "02"
        End If
        jk3.gGlosa1_3n = Trim(TextBox7.Text)

        kp.insertar_qan0300(jk3)

        jk4.gccia_3p = Trim(Label12.Text)
        jk4.gregis_3p = Trim(TextBox1.Text)
        jk4.gOp_3p = Trim(DataGridView1.Rows(0).Cells(0).Value)
        jk4.glinea_3p = Trim(DataGridView1.Rows(0).Cells(1).Value)
        If Trim(DataGridView1.Rows(0).Cells(3).Value) = "" Then
            jk4.ggalga_3p = 0
        Else
            jk4.ggalga_3p = Trim(DataGridView1.Rows(0).Cells(3).Value)
        End If
        If Trim(DataGridView1.Rows(0).Cells(4).Value) = "" Then
            jk4.gAncho_3p = 0
        Else
            jk4.gAncho_3p = Trim(DataGridView1.Rows(0).Cells(4).Value)
        End If
        If Trim(DataGridView1.Rows(0).Cells(5).Value) = "" Then
            jk4.gkgreq_3p = 0
        Else
            jk4.gkgreq_3p = Trim(DataGridView1.Rows(0).Cells(5).Value)
        End If

        jk4.gobserv_3p = ""

        kp.insertar_qanp300(jk4)
        MsgBox("SE REGISTRO LA INFORMACION CON EXITO")
        DataGridView1.Rows.Clear()
        TextBox7.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        TextBox3.Text = ""
        TextBox2.Text = ""
        TextBox7.ReadOnly = True
        TextBox1.Select()
        TY3.Clear()
        ULTIMO()
        abrir()
        llenar_combo_box3()
    End Sub

    Private Sub TextBox1_DoubleClick(sender As Object, e As EventArgs) Handles TextBox1.DoubleClick
        ULTIMO()
        DataGridView1.Rows.Clear()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
    End Sub

    Private Sub TextBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseDown

    End Sub

    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus

    End Sub
End Class