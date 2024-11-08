Imports System.Data.SqlClient

Public Class Rpt_pt
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
	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		abrir()
		Dim cmd As New SqlDataAdapter("select nombfich_3 AS CLIENTE,Q.descri_3 AS DESCRIPCION,sum(cantk_3a) AS CANTIDAD,pedido_3a AS OP,q.program_3 AS PO  from custom_vianny.dbo.fag0400 f inner join custom_vianny.dbo.fap0400 p on f.CIA_3 = p.CIA_3A and f.sfactu_3 =p.sfactu_3a and f.nfactu_3 =p.nfactu_3a 
	  LEFT JOIN  custom_vianny.DBO.qag0300 q on p.pedido_3a = q.ncom_3 and p.CIA_3A = q.ccia
		where almreg_3 ='01' and CIA_3A ='01' and    f.flag_3 ='1' and (f.fcom_3>='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' and f.fcom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "') group by q.program_3,pedido_3a,linea_3a,nombfich_3,Q.descri_3", conx)
		Dim da As New DataTable
		cmd.Fill(da)
		DataGridView1.DataSource = da
		Dim ext As New Exportar
		ext.llenarExcel(DataGridView1)



	End Sub
End Class