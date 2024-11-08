Imports System.Data.SqlClient
Public Class Detalle_guia
    Dim dt As New DataTable
    Public conx As SqlConnection
    Dim Rsr3, Rsr4 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Detalle_guia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hi As New fguiasistema
        Dim jk As New vguiacabecera

        jk.gserie = TextBox2.Text
        jk.gfini = DateSerial(Year(DateTime.Now.ToString("dd/MM/yyyy")), Month(DateTime.Now.ToString("dd/MM/yyyy")), 1)
        jk.gffin = DateSerial(Year(DateTime.Now.ToString("dd/MM/yyyy")), Month(DateTime.Now.ToString("dd/MM/yyyy")) + 1, 0)
        jk.galmacen = Label3.Text
        jk.gccia = Label4.Text

        dt = hi.guia_mostrar(jk)
        If dt.Rows.Count > 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(0).Width = 65
            DataGridView1.Columns(1).Width = 90
            DataGridView1.Columns(2).Width = 240
            DataGridView1.Columns(3).Width = 80
            DataGridView1.Columns(4).Width = 235
            DataGridView1.Columns(5).Width = 210
        End If

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "EMPRESA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 65
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(2).Width = 240
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 235
                DataGridView1.Columns(5).Width = 210
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscarCorrelativo()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CORRELATIVO" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 65
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(2).Width = 240
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 235
                DataGridView1.Columns(5).Width = 210
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim hi As New fguiasistema
        Dim jk As New vguiacabecera

        jk.gserie = TextBox2.Text
        jk.gfini = DateTimePicker1.Value
        jk.gffin = DateTimePicker2.Value
        jk.galmacen = Label3.Text
        jk.gccia = Label4.Text
        dt = hi.guia_mostrar(jk)
        If dt.Rows.Count > 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(0).Width = 65
            DataGridView1.Columns(1).Width = 90
            DataGridView1.Columns(2).Width = 240
            DataGridView1.Columns(3).Width = 80
            DataGridView1.Columns(4).Width = 235
            DataGridView1.Columns(5).Width = 210
        End If
    End Sub





    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscarCorrelativo()
    End Sub
    Sub cabecera(serie, correlativo)
        abrir()
        Dim sql102 As String = "select sfactu_3,nfactu_3,fich_3,isnull(nombfich_3,''),isnull(ppllegad_3,''),isnull(ruc_3,''), isnull(nomb_3,''),isnull(direcc_3,''),isnull(chofer_3,''),isnull(placa_3,''),isnull(brevete_3,''),SUBSTRING(isnull(brevete_3,''),2,8) ,isnull(glosadoc_3,''),trans_3 from  custom_vianny.dbo.Fag0400 where sfactu_3 ='" + serie + "' and nfactu_3 ='" + correlativo + "' and CIA_3 ='" + Label4.Text.ToString().Trim() + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr3 = cmd102.ExecuteReader()
        If Rsr3.Read() = True Then

            Guia_remision_prenda.TextBox45.Text = Rsr3(1).ToString().Trim()
            Guia_remision_prenda.TextBox4.Text = Rsr3(13).ToString().Trim()
            Guia_remision_prenda.TextBox5.Text = Rsr3(6).ToString().Trim()
            Guia_remision_prenda.TextBox6.Text = Rsr3(12).ToString().Trim()
            Guia_remision_prenda.TextBox8.Text = Rsr3(2).ToString().Trim()
            Guia_remision_prenda.TextBox9.Text = Rsr3(3).ToString().Trim()
            Guia_remision_prenda.TextBox10.Text = Rsr3(4).ToString().Trim()
            Guia_remision_prenda.TextBox11.Text = Rsr3(5).ToString().Trim()
            Guia_remision_prenda.TextBox12.Text = Rsr3(6).ToString().Trim()
            Guia_remision_prenda.TextBox13.Text = Rsr3(7).ToString().Trim()
            Guia_remision_prenda.TextBox14.Text = Rsr3(9).ToString().Trim()
            Guia_remision_prenda.TextBox15.Text = Rsr3(10).ToString().Trim()
            Guia_remision_prenda.TextBox16.Text = Rsr3(8).ToString().Trim()
        End If

        Rsr3.Close()
    End Sub
    Sub detalle(serie1, correlativo1)
        abrir()
        Dim i5 As Int32 = 0
        Dim sql102 As String = "SELECT A.*, 
				   ISNULL(B.NOMB_17,'') AS Nomb_17,
				   ISNULL(C.NOMB_sin,B.NOMB_17) AS Nomb_sin,
				   ISNULL(B.IdAlterno_17,'') AS IdAlterno_17,
				   ISNULL(B.Tallaje_17,'') AS Tallaje_17,
				   ISNULL(B.Calidad_17,'') AS Calidad_17,
				   ISNULL(B.TipCtrl_17,'') AS TipCtrl_17,
				   ISNULL(B.Unid_17,'') AS Unid_17,
				    cast( ISNULL(B.Stock_17,0) as int) AS Stock_17,
				  cast( ISNULL(D.Stock1_la,0) as int) AS Stock1_la,
				  cast( ISNULL(D.Stock2_la,0) as int) AS Stock2_la,
				  cast( ISNULL(D.Stock3_la,0) as int) AS Stock3_la,
				  cast( ISNULL(D.Stock4_la,0) as int) AS Stock4_la,
				  cast( ISNULL(D.Stock5_la,0) as int) AS Stock5_la,
				  cast( ISNULL(D.Stock6_la,0) as int) AS Stock6_la,
				  cast( ISNULL(D.Stock7_la,0) as int) AS Stock7_la,
				  cast( ISNULL(D.Stock8_la,0) as int) AS Stock8_la,
				  cast( ISNULL(D.Stock9_la,0) as int) AS Stock9_la,
				  cast( ISNULL(D.Stock10_la,0) as int) AS Stock10_la,
				  cast( ISNULL(Z.Stock1_st,0) as int) AS Stock1_st,
				  cast( ISNULL(Z.Stock2_st,0) as int) AS Stock2_st,
				  cast( ISNULL(Z.Stock3_st,0) as int) AS Stock3_st,
				  cast( ISNULL(Z.Stock4_st,0) as int) AS Stock4_st,
				  cast( ISNULL(Z.Stock5_st,0) as int) AS Stock5_st,
				  cast( ISNULL(Z.Stock6_st,0) as int) AS Stock6_st,
				  cast( ISNULL(Z.Stock7_st,0) as int) AS Stock7_st,
				  cast( ISNULL(Z.Stock8_st,0) as int) AS Stock8_st,
				  cast( ISNULL(Z.Stock9_st,0) as int) AS Stock9_st,
				  cast( ISNULL(Z.Stock10_st,0) as int) AS Stock10_st,
				   ISNULL(F.Talla1_tl,'') AS Talla1_tl,
				   ISNULL(F.Talla2_tl,'') AS Talla2_tl,
				   ISNULL(F.Talla3_tl,'') AS Talla3_tl,
				   ISNULL(F.Talla4_tl,'') AS Talla4_tl,
				   ISNULL(F.Talla5_tl,'') AS Talla5_tl,
				   ISNULL(F.Talla6_tl,'') AS Talla6_tl,
				   ISNULL(F.Talla7_tl,'') AS Talla7_tl,
				   ISNULL(F.Talla8_tl,'') AS Talla8_tl,
				   ISNULL(F.Talla9_tl,'') AS Talla9_tl,
				   ISNULL(F.Talla10_tl,'') AS Talla10_tl,
				   ISNULL(G.Dele,'') AS NTalla1,
	       		   ISNULL(H.Dele,'') AS NTalla2,
	       		   ISNULL(I.Dele,'') AS NTalla3,
	       		   ISNULL(J.Dele,'') AS NTalla4,
	       		   ISNULL(K.Dele,'') AS NTalla5,
	       		   ISNULL(L.Dele,'') AS NTalla6,
	       		   ISNULL(M.Dele,'') AS NTalla7,
	       		   ISNULL(N.Dele,'') AS NTalla8,
	       		   ISNULL(P.Dele,'') AS NTalla9,
	       		   ISNULL(Q.Dele,'') AS NTalla10,
	       		   ISNULL(R.Dele,'') AS NMarca,
	       		   ISNULL(R.Nele,0) AS TRubro,
	       		   ISNULL(S.DesComb,'') AS NColor,
	       		   ISNULL(T.Dele,'') AS NLavado
				   FROM custom_vianny.dbo.Fam0400 A LEFT JOIN custom_vianny.dbo.Fag0400 U
				   		ON A.CCia_3m = U.Cia_3 AND A.SFactu_3m = U.SFactu_3 AND A.NFactu_3m = U.NFactu_3
				   		LEFT JOIN custom_vianny.dbo.CAG1700 B 
				   		ON A.CCia_3m = B.CCIA AND A.LINEA_3m = B.linea_17
				   		LEFT JOIN custom_vianny.dbo.Cag1700_Sinonimos C
				   		ON A.CCia_3m = C.CCia_Sin AND A.Linea_3m = C.Codigo_sin AND A.Sinon_3m = C.Item_sin
						LEFT JOIN custom_vianny.dbo.Cag1700_Almac_Lotes D
				   		ON A.CCia_3m = D.CCia_la AND U.AlmReg_3 = D.Almac_la AND A.Linea_3m = D.Codigo_la AND A.Pedido_3m = D.Lote_la
				   		LEFT JOIN custom_vianny.dbo.Tallado F
				   		ON B.CCia = F.CCia_tl AND B.Tallaje_17 = F.Codigo_tl
				   		LEFT JOIN custom_vianny.dbo.Tab0100 G
				   		ON F.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND F.Talla1_tl = LEFT(G.Cele,4)
				   		LEFT JOIN custom_vianny.dbo.Tab0100 H
				   		ON F.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND F.Talla2_tl = LEFT(H.Cele,4)
				   		LEFT JOIN custom_vianny.dbo.Tab0100 I
				   		ON F.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND F.Talla3_tl = LEFT(I.Cele,4)
				   		LEFT JOIN custom_vianny.dbo.Tab0100 J
				   		ON F.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND F.Talla4_tl = LEFT(J.Cele,4)
				   		LEFT JOIN custom_vianny.dbo.Tab0100 K
				   		ON F.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND F.Talla5_tl = LEFT(K.Cele,4)
				   		LEFT JOIN custom_vianny.dbo.Tab0100 L
				   		ON F.CCia_tl = L.CCia AND 'TALLAS' = L.CTab AND F.Talla6_tl = LEFT(L.Cele,4)
				   		LEFT JOIN custom_vianny.dbo.Tab0100 M
				   		ON F.CCia_tl = M.CCia AND 'TALLAS' = M.CTab AND F.Talla7_tl = LEFT(M.Cele,4)
				   		LEFT JOIN custom_vianny.dbo.Tab0100 N
				   		ON F.CCia_tl = N.CCia AND 'TALLAS' = N.CTab AND F.Talla8_tl = LEFT(N.Cele,4)
				   		LEFT JOIN custom_vianny.dbo.Tab0100 P
				   		ON F.CCia_tl = P.CCia AND 'TALLAS' = P.CTab AND F.Talla9_tl = LEFT(P.Cele,4)
				   		LEFT JOIN custom_vianny.dbo.Tab0100 Q
				   		ON F.CCia_tl = Q.CCia AND 'TALLAS' = Q.CTab AND F.Talla10_tl = LEFT(Q.Cele,4)
			       		LEFT JOIN custom_vianny.dbo.Tab0100 R
		       		  	ON B.CCia = R.CCia AND 'MARCLI' = R.CTab AND B.Marcap_17 = LEFT(R.Cele,4)
						LEFT JOIN custom_vianny.dbo.QagCombinaciones S
    		  		    ON B.CCia = S.CCia AND B.Combo_17 = S.CodComb
    		  		    LEFT JOIN custom_vianny.dbo.Tab0100 T
				   		ON B.CCia = T.CCia AND 'TIPLAV' = T.CTab AND B.Lavado_17 = LEFT(T.Cele,2)
				   		LEFT JOIN custom_vianny.dbo.Cag1700_Stock Z
    		  		    ON A.CCia_3m = Z.CCia_st AND U.AlmReg_3 = Z.Almac_st AND A.Linea_3m = Z.Codigo_st
			 Where A.ccia_3m = '" + Label4.Text.ToString().Trim() + "' AND A.sfactu_3m = '" + serie1 + "' AND A.nfactu_3m = '" + correlativo1 + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr4 = cmd102.ExecuteReader()
        While Rsr4.Read()
            Guia_remision_prenda.DataGridView1.Rows.Add()
            Guia_remision_prenda.DataGridView1.Rows(i5).Cells(2).Value = Rsr4(6).ToString().Trim()
            Guia_remision_prenda.DataGridView1.Rows(i5).Cells(0).Value = Rsr4(3).ToString().Trim()
            Guia_remision_prenda.DataGridView1.Rows(i5).Cells(1).Value = Rsr4(7).ToString().Trim()
            Guia_remision_prenda.DataGridView1.Rows(i5).Cells(3).Value = Rsr4(5).ToString().Trim()
            Guia_remision_prenda.DataGridView1.Rows(i5).Cells(4).Value = Rsr4(34).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(4).Value = Rsr4(9).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(5).Value = Rsr4(10).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(6).Value = Rsr4(11).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(7).Value = Rsr4(12).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(8).Value = Rsr4(13).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(9).Value = Rsr4(14).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(10).Value = Rsr4(15).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(11).Value = Rsr4(16).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(12).Value = Rsr4(17).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(13).Value = Rsr4(18).ToString().Trim()
            'Guia_remision_prenda.DataGridView1.Rows(i5).Cells(14).Value = Rsr4(8).ToString().Trim()
            Guia_remision_prenda.DataGridView1.Rows(i5).Cells(16).Value = Rsr4(36).ToString().Trim()

            i5 = i5 + 1
        End While
        Rsr4.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Label7.Text = 1 Then
            Guia_remision_prenda.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim
            Guia_remision_prenda.TextBox2.Focus()
            SendKeys.Send("{ENTER}")
        Else
            cabecera(DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim, DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim)
            detalle(DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim, DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim)
            Guia_remision_prenda.ComboBox2.SelectedIndex = 0
        End If

        Me.Close()
    End Sub
End Class