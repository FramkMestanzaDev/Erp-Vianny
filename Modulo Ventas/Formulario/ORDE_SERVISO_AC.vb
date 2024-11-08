﻿Imports System.Data.SqlClient
Public Class ORDE_SERVISO_AC
    Public conx As SqlConnection
    Public comando As SqlCommand
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.Enabled = True
            TextBox1.Select()
        Else
            TextBox1.Enabled = False
            TextBox1.Text = ""
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox2.Enabled = True
            TextBox2.Select()
        Else
            TextBox2.Enabled = False
            TextBox2.Text = ""
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
    Dim da As New DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()
        da.Clear()
        DataGridView1.DataSource = ""
        Dim OP, ruc As String
        If Trim(TextBox1.Text).Length = 0 Then
            OP = "null"
        Else
            OP = Convert.ToString(TextBox1.Text)
        End If
        If Trim(TextBox2.Text).Length = 0 Then
            ruc = "null"
        Else
            ruc = Convert.ToString(TextBox2.Text)
        End If
        Dim cmd As New SqlDataAdapter(" SELECT	A.NCOM_4,A.FICH_4,A.FCOM_4,A.TCAM_4 , A.MONE_4  , ISNULL(E.Nomb_4,'') AS Nomb_4,A.GLOA_4  ,A.GLOB_4  ,A.GLOC_4  ,A.GLOD_4  ,
						ISNULL( C.Nomb_10,SPACE(80)) AS Nomb_10,ISNULL( C.Ruc_10,SPACE(80)) AS Ruc_10,B.Item_4a,B.Linea_4a,ISNULL( D.Nomb_17 ,SPACE(80)) AS Nomb_17,
						ISNULL( D.Unid_17,SPACE(3)) AS Unid_17,B.Cant_4a,B.Pedido_4a,B.PreUn1_4a,B.PreUn2_4a,B.Tot1_4a,B.Tot2_4a,
						ISNULL(D.CCosto_17,'') AS CCosto1,ISNULL(G.Nomb_7k,'')   AS NCosto1,B.CCosto_4a,ISNULL(F.Nomb_7k,'') AS NCosto2,B.Obser1_4a,
						B.Obser2_4a,B.Obser3_4a,B.Obser4_4a,B.Obser5_4a
						FROM custom_vianny.DBO.LAG0400 A  INNER JOIN  custom_vianny.DBO.Lap0400 B
									 ON A.CCia_4 = B.CCia_4a AND A.NCom_4 = B.NCom_4a
										LEFT OUTER JOIN  custom_vianny.DBO.Cag1000 C
									 ON '01' = C.CCia AND A.Fich_4 = C.Fich_10
									 	LEFT OUTER JOIN  custom_vianny.DBO.Cag1700 D
							 		 ON B.CCia_4a = D.CCia AND B.Linea_4a = D.Linea_17
										LEFT OUTER JOIN  custom_vianny.DBO.Qag0400 E
									 ON '01' =  E.CCia AND A.SubFase_4 = E.Fase_4
									    LEFT OUTER JOIN   custom_vianny.DBO.Kag0700 F
									 ON B.CCia_4a = F.CCia AND B.CCosto_4a = F.CCos_7k
									 	LEFT OUTER JOIN   custom_vianny.DBO.Kag0700 G
									 ON D.CCia = G.CCia AND D.CCosto_17 = G.CCos_7k
						 Where CCia_4 = '" + Label3.Text + "' AND FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND Pedido_4a=ISNULL(" + OP + ",Pedido_4a) AND A.fich_4=ISNULL(" + ruc + ",A.fich_4) 
						 Order By A.NCOM_4", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        Dim ext As New Exportar
        ext.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Select Case Trim(TextBox1.Text).Length
            Case "1" : TextBox1.Text = "01" & "0000000" & TextBox1.Text
            Case "2" : TextBox1.Text = "01" & "000000" & TextBox1.Text
            Case "3" : TextBox1.Text = "01" & "00000" & TextBox1.Text
            Case "4" : TextBox1.Text = "01" & "0000" & TextBox1.Text
            Case "5" : TextBox1.Text = "01" & "000" & TextBox1.Text
            Case "6" : TextBox1.Text = "01" & "00" & TextBox1.Text
            Case "7" : TextBox1.Text = "01" & "0" & TextBox1.Text
            Case "8" : TextBox1.Text = "01" & TextBox1.Text
            Case "9" : TextBox1.Text = "0" & TextBox1.Text
        End Select
        rp_os.TextBox1.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        rp_os.TextBox2.Text = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
        rp_os.TextBox3.Text = Label3.Text
        rp_os.TextBox4.Text = Trim(TextBox1.Text)
        rp_os.TextBox5.Text = Trim(TextBox2.Text)
        rp_os.Show()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class