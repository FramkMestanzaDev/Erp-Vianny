Imports System.Data.SqlClient
Public Class Ord_Corte
    Public conx As SqlConnection
    Dim ju1, ju2 As New DataTable
    Private Sub Ord_Corte_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'DataGridView1.Rows.Add(9)
        'DataGridView1.Rows(0).Cells(0).Value = "PROGRAMADO"
        'DataGridView1.Rows(1).Cells(0).Value = "CORTADO"
        'DataGridView1.Rows(2).Cells(0).Value = "X CORTAR"
        'DataGridView1.Rows(3).Cells(0).Value = "PROPOR PROG."
        'DataGridView1.Rows(4).Cells(0).Value = "CAPAS PROG."
        'DataGridView1.Rows(5).Cells(0).Value = "PROPOR LIQ."
        'DataGridView1.Rows(6).Cells(0).Value = "CAPAS LIQ."
        'DataGridView1.Rows(7).Cells(0).Value = "PRENDAS PROG."
        'DataGridView1.Rows(8).Cells(0).Value = "PRENDAS REALES"
        TextBox1.Select()
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
    Dim Rsr21, Rsr211, RS22, RS23, RS24 As SqlDataReader
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Add(9)
            DataGridView1.Rows(0).Cells(0).Value = "PROGRAMADO"
            DataGridView1.Rows(1).Cells(0).Value = "CORTADO"
            DataGridView1.Rows(2).Cells(0).Value = "X CORTAR"
            DataGridView1.Rows(3).Cells(0).Value = "PROPOR PROG."
            DataGridView1.Rows(4).Cells(0).Value = "CAPAS PROG."
            DataGridView1.Rows(5).Cells(0).Value = "PROPOR LIQ."
            DataGridView1.Rows(6).Cells(0).Value = "CAPAS LIQ."
            DataGridView1.Rows(7).Cells(0).Value = "PRENDAS PROG."
            DataGridView1.Rows(8).Cells(0).Value = "PRENDAS REALES"
            Select Case TextBox1.Text.Length

                Case "1" : TextBox1.Text = "010000000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "01000000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0100000" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "010000" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = "01000" & "" & TextBox1.Text
                Case "6" : TextBox1.Text = "0100" & "" & TextBox1.Text
                Case "7" : TextBox1.Text = "010" & "" & TextBox1.Text
                Case "8" : TextBox1.Text = "01" & "" & TextBox1.Text
            End Select
            abrir()
            Dim sql1021 As String = "  SELECT c.nomb_10, a.descri_3 FROM custom_vianny.dbo.QAG0300 A left join custom_vianny.dbo.cag1000 c on a.fich_3 = c.fich_10 and a.ccia=c.ccia Where a.ccia = '" + Label44.Text + "' and A.ncom_3 =  '" + TextBox1.Text + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr21 = cmd1021.ExecuteReader()

            If Rsr21.Read() Then
                TextBox2.Text = Rsr21(0)
                TextBox7.Text = Rsr21(1)
            End If

            Rsr21.Close()


            Dim sql10211 As String = "select top 1 ocorte_3cg as id  from custom_vianny.dbo.Qag300Cc where pedido_3cg ='" + TextBox1.Text + "' order by ocorte_3cg desc "
            Dim cmd10211 As New SqlCommand(sql10211, conx)
            Rsr211 = cmd10211.ExecuteReader()

            If Rsr211.Read() Then
                If Trim(Rsr211(0)) = "" Then
                    TextBox3.Text = "001"
                Else
                    TextBox3.Text = Rsr211(0) + 1
                    Select Case TextBox3.Text.Length

                        Case "1" : TextBox3.Text = "00" & "" & TextBox3.Text
                        Case "2" : TextBox3.Text = "0" & "" & TextBox3.Text
                        Case "3" : TextBox3.Text = TextBox3.Text

                    End Select
                End If


            End If

            Rsr21.Close()

            Dim RS2 As SqlDataReader
            Dim KL As New fop
            Dim OL As New vop
            OL.gcia = Label44.Text
            OL.gncom_3 = TextBox1.Text
            abrir()
            Dim sql2 As String = "SELECT COUNT(cant_16d) FROM custom_vianny.dbo.Qag16dv where  ccia_16d = '" + Trim(Label44.Text) + "' And ncom_16d = '" + TextBox1.Text + "' "
            Dim cmd2 As New SqlCommand(sql2, conx)
            RS2 = cmd2.ExecuteReader()
            RS2.Read()
            Dim I1, i2 As Integer

            If RS2(0) <> 0 Then
                ju2 = KL.PADRON_TALLA1(OL)
                DataGridView4.DataSource = ju2
                I1 = DataGridView4.Columns.Count
                i2 = DataGridView4.Rows.Count
                For o = 0 To I1 - 1
                    DataGridView1.Columns.Add(DataGridView4.Rows(0).Cells(o).Value, DataGridView4.Rows(0).Cells(o).Value)
                    DataGridView1.Columns(o + 1).HeaderText = DataGridView4.Columns(o).HeaderText
                    DataGridView1.Columns(o + 1).Name = DataGridView4.Columns(o).HeaderText
                    'DataGridView1.Columns(o).HeaderText.ToString.Trim = DataGridView4.Columns(o).HeaderText.ToString.Trim
                    DataGridView1.Rows(0).Cells(o + 1).Value = DataGridView4.Rows(0).Cells(o).Value
                Next



            Else
                DataGridView4.DataSource = ju1
            End If
            RS2.Close()

            TextBox3.Select()
        End If

    End Sub


    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox3.Text.Length

                Case "1" : TextBox3.Text = "00" & "" & TextBox3.Text
                Case "2" : TextBox3.Text = "0" & "" & TextBox3.Text
                Case "3" : TextBox3.Text = TextBox3.Text

            End Select
            TextBox4.Select()
        End If
    End Sub


    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox4.Text.Length

                Case "1" : TextBox4.Text = "0" & "" & TextBox4.Text
                Case "2" : TextBox4.Text = TextBox4.Text


            End Select
            TextBox1.Enabled = False
            TextBox3.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Select()

        End If
    End Sub


    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox5.Text.Length

                Case "1" : TextBox5.Text = "0" & "" & TextBox5.Text
                Case "2" : TextBox5.Text = TextBox5.Text


            End Select
            abrir()
            Dim sql2 As String = "
					   SELECT A.*,
			          ISNULL(C.Fich_3,'') AS Fich_3,
				      ISNULL(C.FinPed_3,0) AS FinPed_3,
		       		  ISNULL(B.Nomb_17,'') AS Nomb_17,
		       		  ISNULL(D.Sigla_10,'') AS Sigla_10,
		       		  ISNULL(D.Nomb_10,'') AS Nomb_10
		       		  FROM custom_vianny.dbo.Qag300Cc A LEFT JOIN custom_vianny.dbo.Cag1700 B
		       		  	   ON A.Empr_3cg = B.CCia AND A.Tela_3cg = B.Linea_17
		       		  	   			LEFT JOIN custom_vianny.dbo.Qag0300 C
       		  			   ON A.empr_3cg = C.CCia  AND A.pedido_3cg = C.NCom_3
       		  			   			LEFT JOIN custom_vianny.dbo.Cag1000 D
       		  			   ON '01' = D.CCia AND C.Fich_3 = D.Fich_10
		       		   Where A.empr_3cg= '" + Trim(Label44.Text) + "' and A.pedido_3cg ='" + Trim(TextBox1.Text) + "' and A.ocorte_3cg ='" + Trim(TextBox3.Text) + "' and  A.item_3cg='" + Trim(TextBox4.Text) + "' 
		       		   Order By A.Empr_3cg, A.pedido_3cg, a.ocorte_3cg"
            Dim cmd2 As New SqlCommand(sql2, conx)
            RS22 = cmd2.ExecuteReader()
            If RS22.Read() Then
                TextBox9.Text = RS22(9)
                TextBox10.Text = RS22(67)
                TextBox16.Text = RS22(13)
                TextBox35.Text = RS22(14)
                TextBox36.Text = RS22(15)
                TextBox37.Text = RS22(16)
                TextBox36.Text = RS22(17)
                TextBox17.Text = RS22(18)
                TextBox15.Text = RS22(19)
                TextBox20.Text = RS22(20)
                TextBox21.Text = RS22(22)
                TextBox12.Text = RS22(21)
                TextBox14.Text = RS22(23)
                TextBox22.Text = RS22(24)
                TextBox19.Text = RS22(25)
                TextBox31.Text = RS22(26)
                TextBox36.Text = RS22(27)
                TextBox37.Text = RS22(28)
                TextBox38.Text = RS22(29)
                TextBox23.Text = RS22(35)
                TextBox24.Text = RS22(30)
                TextBox18.Text = RS22(31)
                TextBox13.Text = RS22(32)
                TextBox26.Text = RS22(36)
                TextBox27.Text = RS22(37)
                TextBox28.Text = RS22(38)
                TextBox29.Text = RS22(62)
                TextBox30.Text = RS22(64)
                TextBox25.Text = RS22(39)
                DateTimePicker1.Value = RS22(6)
                DateTimePicker2.Value = RS22(7)
                DateTimePicker3.Value = RS22(8)
                If RS22(45) = "1" Then
                    CheckBox1.Checked = True
                Else
                    CheckBox1.Checked = False
                End If
                TextBox9.ReadOnly = True
                TextBox10.ReadOnly = True
                TextBox11.ReadOnly = True
                TextBox12.ReadOnly = True
                TextBox13.ReadOnly = True
                TextBox14.ReadOnly = True
                TextBox15.ReadOnly = True
                TextBox16.ReadOnly = True
                TextBox17.ReadOnly = True
                TextBox18.ReadOnly = True
                TextBox19.ReadOnly = True
                TextBox18.ReadOnly = True
                TextBox20.ReadOnly = True
                TextBox21.ReadOnly = True
                TextBox22.ReadOnly = True
                TextBox23.ReadOnly = True
                TextBox24.ReadOnly = True
                TextBox25.ReadOnly = True
                TextBox26.ReadOnly = True
                TextBox27.ReadOnly = True
                TextBox28.ReadOnly = True
                TextBox29.ReadOnly = True
                TextBox30.ReadOnly = True
                TextBox31.ReadOnly = True
                TextBox32.ReadOnly = True
                TextBox33.ReadOnly = True
                TextBox34.ReadOnly = True
                TextBox35.ReadOnly = True
                TextBox36.ReadOnly = True
                TextBox37.ReadOnly = True

                TextBox38.ReadOnly = True
                DateTimePicker1.Enabled = False
                DateTimePicker2.Enabled = False
                DateTimePicker3.Enabled = False
                CheckBox1.Enabled = False
                DataGridView1.Enabled = False
                DataGridView2.Enabled = False
                DataGridView3.Enabled = False
                TabPage4.Enabled = True

            Else
                TabPage4.Enabled = False
            End If
            RS22.Close()

            Dim sql213 As String = "SELECT A.* FROM custom_vianny.dbo.Qaq300Cc A  Where A.empr_3q = '" + Trim(Label44.Text) + "' and  A.pedido_3q ='" + Trim(TextBox1.Text) + "' and A.corte_3q ='" + Trim(TextBox3.Text) + "' and A.item_3q ='" + Trim(TextBox4.Text) + "' "
            Dim cmd213 As New SqlCommand(sql213, conx)
            RS23 = cmd213.ExecuteReader()
            Dim i51 As Integer
            i51 = 0
            While RS23.Read
                DataGridView2.Rows.Add()
                DataGridView2.Rows(i51).Cells(0).Value = RS23(6)
                DataGridView2.Rows(i51).Cells(1).Value = RS23(8)
                DataGridView2.Rows(i51).Cells(2).Value = RS23(9)
                DataGridView2.Rows(i51).Cells(3).Value = RS23(10)
                DataGridView2.Rows(i51).Cells(4).Value = RS23(11)
                DataGridView2.Rows(i51).Cells(7).Value = RS23(9)
                i51 = i51 + 1
            End While
            RS23.Close()

            Dim sql214 As String = "select  talla_3q,sum(canti_3q) from custom_vianny.dbo.Qaq300cc where pedido_3q  ='" + Trim(TextBox1.Text) + "' and item_3q ='" + Trim(TextBox4.Text) + "' and color_3q ='" + Trim(TextBox5.Text) + "' and corte_3q ='" + Trim(TextBox3.Text) + "'  and Empr_3q ='" + Trim(Label44.Text) + "' group by talla_3q"
            Dim cmd214 As New SqlCommand(sql214, conx)
            RS24 = cmd214.ExecuteReader()
            Dim i54 As Integer
            i54 = 0
            While RS24.Read
                DataGridView3.Rows.Add()
                DataGridView3.Rows(i54).Cells(0).Value = RS24(0)
                DataGridView3.Rows(i54).Cells(1).Value = RS24(1)
                DataGridView3.Rows(i54).Cells(2).Value = RS24(1)

                i54 = i54 + 1
            End While
            RS24.Close()
            TextBox9.Select()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox18.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        TextBox26.Text = ""
        TextBox27.Text = ""
        TextBox28.Text = ""
        TextBox29.Text = ""
        TextBox30.Text = ""
        TextBox31.Text = ""
        TextBox32.Text = ""
        TextBox33.Text = ""
        TextBox34.Text = ""
        TextBox35.Text = ""
        TextBox36.Text = ""
        TextBox37.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        DataGridView1.Rows.Clear()
        Dim p, p1 As Integer
        p = DataGridView1.Columns.Count
        '
        For i = 1 To p - 1
            p1 = DataGridView1.Columns.Count

            'If p1 > 1 Then
            DataGridView1.Columns.RemoveAt(1)
            'End If

        Next

        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
        DataGridView4.DataSource = Nothing
        TextBox1.Enabled = True
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
    End Sub

    Private Sub TextBox1_Move(sender As Object, e As EventArgs) Handles TextBox1.Move

    End Sub
End Class