Imports System.Data.SqlClient
Public Class PadronTalla
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
	Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
		TallasTa.Label1.Text = 1
		TallasTa.ShowDialog()
	End Sub
	Sub mostrarCorrelativo()
		abrir()
		Dim sql10212 As String = "select max(codigo_tl) from custom_vianny.dbo.Tallado"
		Dim cmd10212 As New SqlCommand(sql10212, conx)
		Rsr20 = cmd10212.ExecuteReader()
		If Rsr20.Read() = True Then
			TextBox1.Text = Convert.ToInt64(Rsr20(0)) + 1
		End If

	End Sub
	Public Sub MOSTRAR_INFORMACION()
		abrir()
		DataGridView1.DataSource = ""
		DT14.Clear()
		Dim cmd6 As New SqlDataAdapter(" SELECT A.*,
		       		  ISNULL(B.Dele,'') AS NTalla1,
		       		  ISNULL(C.Dele,'') AS NTalla2,
		       		  ISNULL(D.Dele,'') AS NTalla3,
		       		  ISNULL(E.Dele,'') AS NTalla4,
		       		  ISNULL(F.Dele,'') AS NTalla5,
		       		  ISNULL(G.Dele,'') AS NTalla6,
		       		  ISNULL(H.Dele,'') AS NTalla7,
		       		  ISNULL(I.Dele,'') AS NTalla8,
		       		  ISNULL(J.Dele,'') AS NTalla9,
		       		  ISNULL(K.Dele,'') AS NTalla10
				      FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
				      				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 C
				      				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 D
				      				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 E
				      				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 F
				      				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 G
				      				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 H
				      				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 I
				      				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 J
				      				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
				      				 LEFT JOIN custom_vianny.dbo.Tab0100 K
				      				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
		        Where A.CCia_tl = '01'
		       ORDER BY  A.CCia_tl, A.Codigo_tl", conx)
		cmd6.Fill(DT14)
		DataGridView1.DataSource = DT14
		DataGridView1.Columns(0).Visible = False
		DataGridView1.Columns(2).Visible = False
		DataGridView1.Columns(3).Visible = False
		DataGridView1.Columns(4).Visible = False
		DataGridView1.Columns(5).Visible = False
		DataGridView1.Columns(6).Visible = False
		DataGridView1.Columns(7).Visible = False
		DataGridView1.Columns(8).Visible = False
		DataGridView1.Columns(9).Visible = False
		DataGridView1.Columns(10).Visible = False
		DataGridView1.Columns(11).Visible = False
		DataGridView1.Columns(12).Visible = False
		DataGridView1.Columns(1).Width = 55
		DataGridView1.Columns(13).Width = 55
		DataGridView1.Columns(14).Width = 55
		DataGridView1.Columns(15).Width = 55
		DataGridView1.Columns(16).Width = 55
		DataGridView1.Columns(17).Width = 55
		DataGridView1.Columns(18).Width = 55
		DataGridView1.Columns(19).Width = 55
		DataGridView1.Columns(20).Width = 55
		DataGridView1.Columns(21).Width = 55
		DataGridView1.Columns(22).Width = 55

	End Sub

	Private Sub PadronTalla_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		mostrarCorrelativo()
		MOSTRAR_INFORMACION()
	End Sub

	Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
		TallasTa.Label1.Text = 2
		TallasTa.ShowDialog()
	End Sub

	Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
		TallasTa.Label1.Text = 3
		TallasTa.ShowDialog()
	End Sub

	Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
		TallasTa.Label1.Text = 4
		TallasTa.ShowDialog()
	End Sub

	Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
		TallasTa.Label1.Text = 5
		TallasTa.ShowDialog()

	End Sub

	Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
		TallasTa.Label1.Text = 6
		TallasTa.ShowDialog()
	End Sub

	Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
		TallasTa.Label1.Text = 7
		TallasTa.ShowDialog()
	End Sub

	Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
		TallasTa.Label1.Text = 8
		TallasTa.ShowDialog()
	End Sub

	Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
		TallasTa.Label1.Text = 9
		TallasTa.ShowDialog()
	End Sub

	Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
		TallasTa.Label1.Text = 10
		TallasTa.ShowDialog()
	End Sub

	Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
		Dim cmd20 As New SqlCommand("insert into custom_vianny.dbo.Tallado (ccia_tl,codigo_tl,talla1_tl,talla2_tl,talla3_tl,talla4_tl,talla5_tl,talla6_tl,talla7_tl,talla8_tl,talla9_tl,talla10_tl) values ('01',@codigo_tl,@talla1_tl,@talla2_tl,@talla3_tl,@talla4_tl,@talla5_tl,@talla6_tl,@talla7_tl,@talla8_tl,@talla9_tl,@talla10_tl)", conx)
		cmd20.Parameters.AddWithValue("@codigo_tl", Trim(TextBox1.Text))
		cmd20.Parameters.AddWithValue("@talla1_tl", Trim(Label13.Text))
		cmd20.Parameters.AddWithValue("@talla2_tl", Trim(Label14.Text))
		cmd20.Parameters.AddWithValue("@talla3_tl", Trim(Label15.Text))
		cmd20.Parameters.AddWithValue("@talla4_tl", Trim(Label16.Text))
		cmd20.Parameters.AddWithValue("@talla5_tl", Trim(Label17.Text))
		cmd20.Parameters.AddWithValue("@talla6_tl", Trim(Label18.Text))
		cmd20.Parameters.AddWithValue("@talla7_tl", Trim(Label19.Text))
		cmd20.Parameters.AddWithValue("@talla8_tl", Trim(Label20.Text))
		cmd20.Parameters.AddWithValue("@talla9_tl", Trim(Label21.Text))
		cmd20.Parameters.AddWithValue("@talla10_tl", Trim(Label22.Text))

		cmd20.ExecuteNonQuery()
		MsgBox(" SE REGISTRO CORRECTAMENTE LA INFORMACION")
		limpiar()
		mostrarCorrelativo()
		MOSTRAR_INFORMACION()
		Form_Padron.mostrar()
	End Sub
	Sub limpiar()

		TextBox2.Text = ""
		TextBox3.Text = ""
		TextBox4.Text = ""
		TextBox5.Text = ""
		TextBox6.Text = ""
		TextBox7.Text = ""
		TextBox8.Text = ""
		TextBox9.Text = ""
		TextBox10.Text = ""
		TextBox11.Text = ""
		Label13.Text = ""
		Label14.Text = ""
		Label15.Text = ""
		Label16.Text = ""
		Label17.Text = ""
		Label18.Text = ""
		Label19.Text = ""
		Label20.Text = ""
		Label21.Text = ""
		Label22.Text = ""

	End Sub
	Sub BUSCAR()
		Dim ds As New DataSet
		ds.Tables.Add(DT14.Copy)
		Dim dv As New DataView(ds.Tables(0))
		dv.RowFilter = "NTalla1" & " like '%" & TextBox12.Text & "%'"
		If dv.Count <> 0 Then
			DataGridView1.DataSource = dv
			DataGridView1.Columns(0).Width = 110
			DataGridView1.Columns(1).Width = 340
		Else
			DataGridView1.DataSource = Nothing
		End If
	End Sub
	Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
		BUSCAR()
	End Sub

	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

		CrearTalla.ShowDialog()
	End Sub

	Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
		If e.KeyCode = Keys.F1 Then
			TallasTa.Label1.Text = 1
			TallasTa.TextBox2.Text = ""

			TallasTa.ShowDialog()
		End If
	End Sub

	Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
		If e.KeyCode = Keys.F1 Then
			TallasTa.Label1.Text = 2
			TallasTa.TextBox2.Text = ""

			TallasTa.ShowDialog()
		End If
	End Sub

	Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
		If e.KeyCode = Keys.F1 Then
			TallasTa.Label1.Text = 3
			TallasTa.TextBox2.Text = ""
			TallasTa.ShowDialog()
		End If
	End Sub

	Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
		If e.KeyCode = Keys.F1 Then
			TallasTa.Label1.Text = 4
			TallasTa.TextBox2.Text = ""
			TallasTa.ShowDialog()
		End If
	End Sub

	Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
		If e.KeyCode = Keys.F1 Then
			TallasTa.Label1.Text = 5
			TallasTa.TextBox2.Text = ""
			TallasTa.ShowDialog()
		End If
	End Sub

	Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
		If e.KeyCode = Keys.F1 Then
			TallasTa.Label1.Text = 6
			TallasTa.TextBox2.Text = ""
			TallasTa.ShowDialog()
		End If
	End Sub

	Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
		If e.KeyCode = Keys.F1 Then
			TallasTa.Label1.Text = 7
			TallasTa.TextBox2.Text = ""
			TallasTa.ShowDialog()
		End If
	End Sub

	Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
		If e.KeyCode = Keys.F1 Then
			TallasTa.Label1.Text = 8
			TallasTa.TextBox2.Text = ""
			TallasTa.ShowDialog()
		End If
	End Sub

	Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
		If e.KeyCode = Keys.F1 Then
			TallasTa.Label1.Text = 9
			TallasTa.TextBox2.Text = ""
			TallasTa.ShowDialog()
		End If
	End Sub

	Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
		If e.KeyCode = Keys.F1 Then
			TallasTa.Label1.Text = 10
			TallasTa.TextBox2.Text = ""
			TallasTa.ShowDialog()
		End If
	End Sub
End Class