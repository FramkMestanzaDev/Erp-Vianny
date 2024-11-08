Imports System.Data.SqlClient
Public Class ReporteSalidas
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
		Dim almacen As String
		Select Case Trim(ComboBox1.Text)
			Case "PRODUCTO TERMINADO" : almacen = "01"
			Case "AVIOS" : almacen = "13"
			Case "TELA PLANA" : almacen = "08"
			Case "TELA PUNTO" : almacen = "10"
			Case "HILO" : almacen = "03"
		End Select

		Dim cmd As New SqlDataAdapter("EXEC custom_vianny.DBO.caeSoft_ReportePartedeIngresoSalida '" + Trim(Label4.Text) + "','" + almacen + "','2','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "','" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "',NULL,NULL,NULL,NULL,NULL,NULL,2,4", conx)
		Dim da As New DataTable
		cmd.Fill(da)
		DataGridView1.DataSource = da
		Dim ext As New Exportar
		ext.llenarExcel(DataGridView1)
	End Sub

	Private Sub ReporteSalidas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		ComboBox1.SelectedIndex = 0
	End Sub
End Class