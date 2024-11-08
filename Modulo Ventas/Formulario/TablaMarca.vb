Imports System.Data.SqlClient
Public Class TablaMarca
	Public conx As SqlConnection
	Public comando As SqlCommand
	Public conn As SqlDataAdapter
	Dim Rsr20 As SqlDataReader
	Sub abrir()
		Try
			conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
			'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

			conx.Open()

		Catch ex As Exception
			MsgBox(ex.Message)

		End Try
	End Sub
	Dim DT14 As New DataTable
	Private Sub TablaMarca_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		mostrarCorrelativo()
		MOSTRAR_INFORMACION()
	End Sub
	Public Sub MOSTRAR_INFORMACION()
		abrir()
		DataGridView1.DataSource = ""
		DT14.Clear()
		Dim cmd6 As New SqlDataAdapter("SELECT cele as CODIGO,dele AS MARCA,codigo AS ABREVIATURA FROM custom_vianny.dbo.TAB0100   Where ccia + ctab='01MARCLI' and  codigo2 ='1' ORDER by  cast(cele as integer)", conx)
		cmd6.Fill(DT14)
		DataGridView1.DataSource = DT14
		DataGridView1.Columns(2).Width = 220
		DataGridView1.Columns(1).Width = 220
		DataGridView1.Columns(0).Width = 85

	End Sub
	Sub mostrarCorrelativo()
		abrir()
		Dim sql10212 As String = "select MAX(cele)  FROM custom_vianny.dbo.TAB0100   Where ccia + ctab='01MARCLI'"
		Dim cmd10212 As New SqlCommand(sql10212, conx)
		Rsr20 = cmd10212.ExecuteReader()
		If Rsr20.Read() = True Then
			TextBox1.Text = Convert.ToInt64(Rsr20(0)) + 1
		End If

	End Sub

	Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
		Dim cmd20 As New SqlCommand("insert into custom_vianny.dbo.TAB0100 (ccia,ctab,cele,dele,nele,codigo,codigo2) values ('01','MARCLI',@cele,@dele,0,@codigo,'1')", conx)
		cmd20.Parameters.AddWithValue("@cele", Trim(TextBox1.Text))
		cmd20.Parameters.AddWithValue("@dele", Trim(TextBox2.Text))
		cmd20.Parameters.AddWithValue("@codigo", Trim(TextBox3.Text))
		cmd20.ExecuteNonQuery()
		MsgBox(" SE REGISTRO CORRECTAMENTE LA INFORMACION")
		limpiar()
		mostrarCorrelativo()
		MOSTRAR_INFORMACION()
	End Sub
	Sub limpiar()
		TextBox2.Text = ""
		TextBox3.Text = ""
	End Sub
End Class