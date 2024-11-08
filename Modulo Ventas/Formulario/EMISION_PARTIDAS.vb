Imports System.Data.SqlClient
Public Class EMISION_PARTIDAS
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim TY4 As New DataTable
    Public conn As SqlDataAdapter
    Sub abrir()
        Try

            conx = New SqlConnection("Data Source=172.21.0.1;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Det_Rollo.Label1.Text = "PO"
            Det_Rollo.Label2.Text = 7
            Det_Rollo.Show()

        Else
            If e.KeyCode = Keys.Enter Then
                MOSTRAR()
            End If

        End If
    End Sub
    Dim ko As New DataTable
    Sub MOSTRAR()
        DataGridView4.DataSource = ""
        DataGridView1.Rows.Clear()
        Dim fj As New fprogramacion
        Dim jk As New vprogramacion
        Dim Rs, Rs1 As SqlDataReader
        jk.gccia = Trim(Label2.Text)
        jk.gop = TextBox1.Text
        ko = fj.buscar_ot_partida(jk)
        If ko.Rows.Count <> 0 Then
            Dim p As Integer

            DataGridView4.DataSource = ko
            p = DataGridView4.Rows.Count
            DataGridView1.Rows.Add(p)
            For i = 0 To p - 1
                DataGridView1.Rows(i).Cells(1).Value = DataGridView4.Rows(i).Cells(1).Value
                DataGridView1.Rows(i).Cells(2).Value = DataGridView4.Rows(i).Cells(6).Value + " - " + DataGridView4.Rows(i).Cells(7).Value
                DataGridView1.Rows(i).Cells(3).Value = DataGridView4.Rows(i).Cells(3).Value
                DataGridView1.Rows(i).Cells(4).Value = DataGridView4.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(5).Value = DataGridView4.Rows(i).Cells(8).Value
                DataGridView1.Rows(i).Cells(10).Value = DataGridView4.Rows(i).Cells(2).Value
                abrir()
                Dim sql1 As String = "select isnull(SUM(kgreq_3p),0) from  custom_vianny.dbo.Qanp300 where Op_3p ='" + Trim(DataGridView4.Rows(i).Cells(1).Value) + "' and ccia_3p='" + Label2.Text + "'"
                Dim cmd1 As New SqlCommand(sql1, conx)
                Rs1 = cmd1.ExecuteReader()
                Rs1.Read()
                If Rs1(0) > 0 Then
                    DataGridView1.Rows(i).Cells(6).Value = Rs1(0)

                End If
                Rs1.Close()
                Dim sql10 As String = "select  isnull(SUM(m.cantkmv_3r),0) from custom_vianny.dbo.marg0001 m inner join custom_vianny.dbo.Qanp300 q on m.partida_3r = q.regis_3p where  flag_3r > 1 AND Q.Op_3p ='" + Trim(DataGridView4.Rows(i).Cells(1).Value) + "' and ccia_3p='" + Label2.Text + "' "
                Dim cmd10 As New SqlCommand(sql10, conx)
                Rs = cmd10.ExecuteReader()
                Rs.Read()
                If Rs(0) > 0 Then
                    DataGridView1.Rows(i).Cells(7).Value = Rs(0)

                End If
                Rs.Close()
                DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(5).Value - DataGridView1.Rows(i).Cells(6).Value
            Next
        Else
            DataGridView4.DataSource = ""
        End If
        TextBox1.Enabled = False
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        MOSTRAR()
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        DataGridView1.BeginEdit(True)
        If e.ColumnIndex = 0 Then


            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            If DataGridView1.Rows(num1).Cells(0).Value = False Then
                DataGridView1.Rows(num1).Cells(0).Value = True
                DataGridView1.Rows(num1).Cells(9).ReadOnly = False

                DataGridView1.Rows(num1).Cells(9).Selected = True
                Button6.Enabled = True

            Else
                If DataGridView1.Rows(num1).Cells(0).Value = True Then

                    DataGridView1.Rows(num1).Cells(0).Value = False
                    DataGridView1.Rows(num1).Cells(9).Value = "0.00"
                    DataGridView1.Rows(num1).Cells(9).ReadOnly = True
                End If

            End If
        End If
        If e.ColumnIndex = 9 Then
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            If DataGridView1.Rows(num1).Cells(9).Value > DataGridView1.Rows(num1).Cells(8).Value Then
                MsgBox("Estado_Cliente PROGRAMANDO MAS DE LO SOLICITADO")
            End If
        End If

    End Sub


    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged

        Dim J As Integer

        J = DataGridView1.Rows.Count
        For I = 0 To J - 1
            If DataGridView1.CurrentRow.Index.ToString() = I Then
                'MsgBox(I)

                DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Yellow

            Else

                    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.White
            End If
        Next
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        DataGridView1.BeginEdit(True)
        abrir()
        DataGridView3.DataSource = ""
        DataGridView2.DataSource = ""
        Dim dg1 As New DataTable
        Dim cmd As New SqlDataAdapter("select regis_3p as PARTIDA,QG.fcom_3n AS FECHA,c.nomb_17 AS CODIGO, QP.kgreq_3p  from custom_vianny.dbo.Qanp300 QP 
					   INNER JOIN custom_vianny.dbo.qan0300 QG ON QP.ccia_3p = QP.ccia_3p AND QP.regis_3p =QG.regis_3n 
					   LEFT JOIN custom_vianny.dbo.cag1700 c on qp.ccia_3p = c.ccia and qp.linea_3p = c.linea_17
					   where Op_3p ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) + "' and qg.flag_3n =1  and qp.ccia_3p ='" + Label2.Text + "' order by regis_3p ", conx)

        cmd.Fill(dg1)
        If dg1.Rows.Count <> 0 Then
            DataGridView3.DataSource = dg1
            DataGridView3.Columns(2).Width = 550
        Else
            DataGridView3.DataSource = ""
        End If

        Dim dg2 As New DataTable
        Dim cmd2 As New SqlDataAdapter("Select ncom_4a as OT,G.partidA_4 AS PARTIDA,G.fcom_4 AS FECHA,P.cantreq_4a AS KGS_PROGRAMADO,P.maquina_4a AS MAQ,P.diamet_4a AS D,P.galga_4a G,P.ancho_4a AS A,P.lm1_41a AS LM1, P.lm2_41a AS LM2,P.lm3_41a AS LM2, P.tela_4a  
			 from  custom_vianny.dbo.matp040f P INNER JOIN custom_vianny.dbo.matg040f G ON P.ncom_4a = G.ncom_4 AND G.ccia_4 = P.ccia_4a 
			 where SUBSTRING(ncom_4a,1,10) ='" + Trim(TextBox1.Text) + "' AND G.flag_4=1 and p.ccia_4a ='" + Label2.Text + "' order by partidA_4", conx)

        cmd2.Fill(dg2)
        If dg2.Rows.Count <> 0 Then
            DataGridView2.DataSource = dg2
            'DataGridView3.Columns(2).Width = 350
            DataGridView2.Columns(11).Visible = False
        Else
            DataGridView2.DataSource = ""
        End If
        Label6.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox1.Enabled = True
        DataGridView2.DataSource = ""
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim Rsr, Rsr1 As SqlDataReader
        abrir()
        Dim sql100 As String = "EXEC custom_vianny.dbo.CAESOFT_GetAllUltimaVersionOrdenTejeduria '" + Label2.Text + "','" + Trim(TextBox1.Text) + " ',NULL"
        Dim cmd100 As New SqlCommand(sql100, conx)
        Rsr = cmd100.ExecuteReader()
        Rsr.Read()
        Otejeduria.TextBox1.Text = TextBox1.Text
        Otejeduria.TextBox2.Text = Rsr(0)
        Rsr.Close()
        Dim sql101 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAllUltimaPartida '" + Label2.Text + "','" + Trim(Label3.Text) + "'"
        Dim cmd101 As New SqlCommand(sql101, conx)
        Rsr1 = cmd101.ExecuteReader()
        Rsr1.Read()
        Otejeduria.TextBox5.Text = Rsr1(0)
        Rsr1.Close()
        Otejeduria.Label14.Text = Label2.Text
        Otejeduria.Label13.Text = Label3.Text
        Otejeduria.Label16.Text = Label4.Text
        Otejeduria.Label17.Text = 1
        Dim p As Integer
        p = DataGridView1.Rows.Count
        For i = 0 To p - 1
            If DataGridView1.Rows(i).Cells(0).Value = True Then
                If DataGridView1.Rows(i).Cells(9).Value <> "" Then
                    Otejeduria.DataGridView1.Rows.Add()
                    Otejeduria.DataGridView1.Rows(0).Cells(0).Value = DataGridView1.Rows(i).Cells(10).Value
                    Otejeduria.DataGridView1.Rows(0).Cells(10).Value = DataGridView1.Rows(i).Cells(9).Value
                    Otejeduria.TextBox6.Text = Mid(DataGridView1.Rows(i).Cells(2).Value, 1, 6)
                    Otejeduria.TextBox7.Text = Trim(Mid(DataGridView1.Rows(i).Cells(2).Value, 9, 20))
                    Otejeduria.TextBox12.Text = Trim(DataGridView1.Rows(i).Cells(1).Value)
                    Otejeduria.Show()
                    Button6.Enabled = False

                Else
                    MsgBox("DEBE INGRESAR UNA CANTIDAD A PORGRAMAR EN LA OP SELECCIONADA")
                    Otejeduria.DataGridView2.Rows.Clear()
                    Otejeduria.DataGridView1.Rows.Clear()
                End If

            End If

        Next



    End Sub

    Private Sub EMISION_PARTIDAS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Otejeduria.DataGridView1.Rows.Add()
        Otejeduria.TextBox1.Text = Mid(DataGridView2.Rows(Trim(Label5.Text)).Cells(0).Value, 1, 10)
        Otejeduria.TextBox2.Text = Mid(DataGridView2.Rows(Trim(Label5.Text)).Cells(0).Value, 11, 3)
        Otejeduria.DataGridView1.Rows(0).Cells(0).Value = DataGridView2.Rows(Trim(Label5.Text)).Cells(11).Value
        Otejeduria.DataGridView1.Rows(0).Cells(1).Value = DataGridView2.Rows(Trim(Label5.Text)).Cells(4).Value
        Otejeduria.DataGridView1.Rows(0).Cells(2).Value = DataGridView2.Rows(Trim(Label5.Text)).Cells(5).Value
        Otejeduria.DataGridView1.Rows(0).Cells(3).Value = DataGridView2.Rows(Trim(Label5.Text)).Cells(6).Value
        Otejeduria.DataGridView1.Rows(0).Cells(4).Value = DataGridView2.Rows(Trim(Label5.Text)).Cells(7).Value
        Otejeduria.DataGridView1.Rows(0).Cells(5).Value = DataGridView2.Rows(Trim(Label5.Text)).Cells(8).Value
        Otejeduria.DataGridView1.Rows(0).Cells(6).Value = DataGridView2.Rows(Trim(Label5.Text)).Cells(9).Value
        Otejeduria.DataGridView1.Rows(0).Cells(7).Value = DataGridView2.Rows(Trim(Label5.Text)).Cells(10).Value
        Otejeduria.TextBox12.Text = Label6.Text
        Otejeduria.Label14.Text = Label2.Text
        Otejeduria.Label13.Text = Label3.Text
        Otejeduria.Label16.Text = Label4.Text
        Otejeduria.Label17.Text = 2
        Otejeduria.Show()
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        Label5.Text = e.RowIndex
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Label5.Text = e.RowIndex
    End Sub
    Dim Rsr2, Rsr25, Rsr26, Rsr27 As SqlDataReader

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        Label7.Text = e.RowIndex
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        abrir()
        Dim sql1022 As String = "select COUNT(regis_3n) from  custom_vianny.DBO.qan0300 where regis_3n ='" + Trim(DataGridView3.Rows(Trim(Label7.Text)).Cells(0).Value) + "' and ccia_3n ='" + Label2.Text + "'"
        Dim cmd1022 As New SqlCommand(sql1022, conx)
        Rsr25 = cmd1022.ExecuteReader()
        Dim i5 As Integer
        i5 = 0
        Rsr25.Read()
        If Rsr25(0) > 0 Then
            CREAR_PARTIDAS.TextBox1.Text = Trim(DataGridView3.Rows(Trim(Label7.Text)).Cells(0).Value)
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
		       		   Where CCia_3n = '" + Label2.Text + "' AND Regis_3n = '" + Trim(DataGridView3.Rows(Trim(Label7.Text)).Cells(0).Value) + "'  
		       		   Order By A.ccia_3N , A.regis_3n"
            Dim cmd1023 As New SqlCommand(sql1023, conx)
            Rsr26 = cmd1023.ExecuteReader()

            Rsr26.Read()
            CREAR_PARTIDAS.TextBox2.Text = Rsr26(10)
            CREAR_PARTIDAS.TextBox3.Text = Rsr26(2)
            CREAR_PARTIDAS.TextBox4.Text = Rsr26(8)
            CREAR_PARTIDAS.TextBox5.Text = Rsr26(4)
            CREAR_PARTIDAS.TextBox6.Text = Rsr26(7)
            CREAR_PARTIDAS.TextBox7.Text = Rsr26(11)
            If Rsr26(5) = "TC" Then
                CREAR_PARTIDAS.ComboBox1.SelectedIndex = 0
            Else
                CREAR_PARTIDAS.ComboBox1.SelectedIndex = 1
            End If
            If Rsr26(6) = "01" Then
                CREAR_PARTIDAS.ComboBox2.SelectedIndex = 0
            Else
                CREAR_PARTIDAS.ComboBox2.SelectedIndex = 1
            End If

            If Rsr26(1) = "INT" Then
                CREAR_PARTIDAS.ComboBox3.SelectedIndex = 0
            Else
                CREAR_PARTIDAS.ComboBox3.SelectedIndex = 1
            End If
            CREAR_PARTIDAS.DateTimePicker1.Value = Rsr26(0)
            CREAR_PARTIDAS.Label13.Text = Rsr26(12)
            Rsr26.Close()
            Dim sql102 As String = "SELECT A.ccia_3p,regis_3p,linea_17,a.galga_3p,a.kgreq_3p,a.ancho_3p,a.Op_3p ,
             ISNULL(B.Nomb_17,'') AS Nomb_17
             FROM custom_vianny.DBO.qanp300 A  LEFT JOIN custom_vianny.DBO.Cag1700 B
            		ON '01' = B.CCia AND A.Linea_3p = B.Linea_17
              Where CCia_3p = '" + Label2.Text + "' AND Regis_3p = '" + Trim(DataGridView3.Rows(Trim(Label7.Text)).Cells(0).Value) + "' 
              Order By A.regis_3p"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr27 = cmd102.ExecuteReader()
            Rsr27.Read()
            CREAR_PARTIDAS.DataGridView1.Rows.Add()
            CREAR_PARTIDAS.DataGridView1.Rows(0).Cells(0).Value = Rsr27(6)
            CREAR_PARTIDAS.DataGridView1.Rows(0).Cells(1).Value = Rsr27(2)
            CREAR_PARTIDAS.DataGridView1.Rows(0).Cells(2).Value = Rsr27(7)
            CREAR_PARTIDAS.DataGridView1.Rows(0).Cells(3).Value = Rsr27(3)
            CREAR_PARTIDAS.DataGridView1.Rows(0).Cells(4).Value = Rsr27(5)
            CREAR_PARTIDAS.DataGridView1.Rows(0).Cells(5).Value = Rsr27(4)
            Rsr27.Close()

            Rsr25.Close()
            abrir()
            llenar_combo_box4(CREAR_PARTIDAS.Label13.Text.ToString)
            Button1.Enabled = False
            CREAR_PARTIDAS.Label14.Text = 1
            CREAR_PARTIDAS.Show()
        Else
            MsgBox("NO HAY INFORMACION PARA MOSTRAR")
        End If

    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Sub llenar_combo_box4(AB As String)
        Try

            conn = New SqlDataAdapter("SELECT fase_4,nomb_4 FROM custom_vianny.dbo.Qag0400 A  Where A.ccia = '" + Label2.Text + "' AND substring(fase_4,1,2)='02' AND fase_4='" + Trim(AB) + "' ", conx)
            conn.Fill(TY4)
            CREAR_PARTIDAS.ComboBox4.DataSource = TY4
            CREAR_PARTIDAS.ComboBox4.DisplayMember = "nomb_4"
            CREAR_PARTIDAS.ComboBox4.ValueMember = "fase_4"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Label7.Text = e.RowIndex
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button6.Select()
        End If
    End Sub
End Class