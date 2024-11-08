Imports System.Data.SqlClient
Public Class Produccion_diaria_cortee
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
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
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
        If Label3.Text = "01" Then
            Prod_diariacorte.TextBox1.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            Prod_diariacorte.TextBox2.Text = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
            Prod_diariacorte.TextBox3.Text = Label3.Text
            Prod_diariacorte.TextBox4.Text = Trim(TextBox1.Text)
            Prod_diariacorte.Show()
        Else
            Rpt_Produccion_Corte_Graus.TextBox1.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            Rpt_Produccion_Corte_Graus.TextBox2.Text = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
            Rpt_Produccion_Corte_Graus.Show()
        End If

    End Sub
    Dim da As New DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()
        da.Clear()
        DataGridView1.DataSource = ""
        Dim OP As String
        If Trim(TextBox1.Text).Length = 0 Then
            OP = "null"
        Else
            OP = Convert.ToString(TextBox1.Text)
        End If
        Dim cmd As New SqlDataAdapter("SELECT A.pedido_3cg,A.ocorte_3cg,A.fdesp_3cg,A.tela_3cg,A.efic_3cg,A.partid_3cg,A.kgprog_3cg,A.conrea_3cg,A.kgreal_3cg,A.kgpart1_3cg,A.prerea_3cg,
			          ISNULL(C.Fich_3,'') AS Fich_3,
				      ISNULL(C.FinPed_3,0) AS FinPed_3,
		       		  ISNULL(B.Nomb_17,'') AS Nomb_17,
		       		  
		       		  ISNULL(D.Nomb_10,'') AS Nomb_10
				 
		       		  FROM custom_vianny.dbo.Qag300Cc A LEFT JOIN custom_vianny.dbo.Cag1700 B
		       		  	   ON A.Empr_3cg = B.CCia AND A.Tela_3cg = B.Linea_17
		       		  	   			LEFT JOIN custom_vianny.dbo.Qag0300 C
       		  			   ON A.empr_3cg = C.CCia  AND A.pedido_3cg = C.NCom_3
       		  			   			LEFT JOIN custom_vianny.dbo.Cag1000 D
       		  			   ON '01' = D.CCia AND C.Fich_3 = D.Fich_10
		       		   Where A.empr_3cg ='01' AND ( A.FDesp_3cg>='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FDesp_3cg<='" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "') AND A.Item_3cg = '01'   and pedido_3cg = ISNULL(" + OP + ", pedido_3cg)
		       		   Order By A.pedido_3cg,A.Empr_3cg,  a.ocorte_3cg  ", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        Dim ext As New Exportar
        ext.llenarExcel(DataGridView1)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Produccion_diaria_cortee_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class