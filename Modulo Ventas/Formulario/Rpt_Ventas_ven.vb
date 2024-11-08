Imports System.Data.SqlClient
Public Class Rpt_Ventas_ven
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
		Dim gmes As String
		Select Case ComboBox1.Text
			Case "ENERO" : gmes = "01"
			Case "FEBRERO" : gmes = "02"
			Case "MARZO" : gmes = "03"
			Case "ABRIL" : gmes = "04"
			Case "MAYO" : gmes = "05"
			Case "JUNIO" : gmes = "06"
			Case "JULIO" : gmes = "07"
			Case "AGOSTO" : gmes = "08"
			Case "SETIEMBRE" : gmes = "09"
			Case "OCTUBRE" : gmes = "10"
			Case "NOVIEMBRE" : gmes = "11"
			Case "DICIEMBRE" : gmes = "12"
			Case "TODOS" : gmes = "NULL"

		End Select
		Dim cmd As New SqlDataAdapter("Set Language spanish  
SELECT a.nomb_3  AS CLIENTE,	  
ISNULL( c.nomb_17,'') AS PRODCUTO,
	 case when 	b.grm_3a = '' then a.sfactur_3+ '-'+a.nfactur_3 else 	b.grm_3a end AS GRM,
	 a.sfactu_3 AS SERIE_FAC,
	 a.nfactu_3 CORRELATIVO_FAC,
	  CASE WHEN A.TiDoc_3 = '007' THEN -1 ELSE 1 END *  B.VVta1_3a   AS VALOR_VENTA,
	 B.preun1_3a  AS PRECIO_S,

	CASE 
	WHEN  a.tienda_3 ='01' THEN 'PRIMERA' 
	WHEN  a.tienda_3 ='37' THEN 'PRIMERA' 
	WHEN  a.tienda_3 ='38' THEN 'SEGUNDA' 
	WHEN  a.tienda_3 ='10' THEN 'PRIMERA' 
	WHEN  a.tienda_3 ='01' THEN 'SEGUNDA' 
	WHEN  a.tienda_3 ='61' THEN 'TERCERA' 
	WHEN  a.tienda_3 ='59' THEN 'SERVICIO' 
	WHEN  a.tienda_3 ='42' THEN 'HILO' 
	WHEN  a.tienda_3 ='06' THEN 'HILO' 
WHEN  a.tienda_3 ='39' THEN 'TERCERA' 
		WHEN  a.tienda_3 ='03' THEN 'HILO' 
		WHEN  a.tienda_3 ='68' THEN 'HUACHIPA' 
ELSE 'VACIO'
	END as ALMACEN,
	

	isnull(k.nomb_15k,'') AS FO_PAGO,
B.cant_3a AS CANTIDAD ,A.fcom_3 AS FECHA
	 FROM custom_vianny.DBO.Fag0300 A LEFT JOIN custom_vianny.DBO.Fap0300 B
					ON A.CCia_3 = B.CCia_3a AND A.CPeriodo_3 = B.CPeriodo_3a AND A.Tienda_3 = B.Tienda_3a AND A.NCom_3 = B.NCom_3a	
					LEFT JOIN custom_vianny.DBO.cag1700 c ON b.ccia_3a=c.ccia and b.linea_3a = c.linea_17
left join custom_vianny.dbo.kag1500 k on  A.condp_3 = k.cond_15k and k.ccia_15k = '" + Trim(Label5.Text) + "'
left join custom_vianny.dbo.Vendedores v on a.vendedor_3 = v.codigo_ven
	 WHERE cperiodo_3 = '" + Trim(TextBox1.Text) + "' AND ccia_3 ='" + Trim(Label5.Text) + "' AND tidoc_3 in (001,003,007)   AND flag_3 ='1' and MONTH(fcom_3)= ISNULL(" + gmes + ", Month(fcom_3)) and A.vendedor_3  = '" + Trim(Label6.Text) + "'

	 UNION ALL

	SELECT nomb_3 AS CLIENTE,


ISNULL(C.nomb_17,'') AS PRODUCTO,
	grm_3a AS GRM,sfactu_3 AS SERIE_FAC,nfactu_3 AS CORRELATIVO_FAC,	rv.tot1_3a AS VALOR_VENTA, preun1_3a AS PRECIO_S,
	'VENTA LIBRE' as ALMACEN,isnull(k.nomb_15k,'')  AS FO_PAGO,cant_3a AS CANTIDAD, VC.fcom_3 AS FECHA from venta_cabecera vc inner join rventa_detalle rv on vc.cperiodo_3 = rv.cperiodo_3a  and vc.empresa = rv.empresa
	and vc.ncom_3 = rv.ncom_3a 
	LEFT JOIN custom_vianny.DBO.cag1700 c ON  RV.linea_3a = c.linea_17
     left join custom_vianny.dbo.kag1500 k on  vc.condp_3 = k.cond_15k and k.ccia_15k = '" + Trim(Label5.Text) + "'
      left join custom_vianny.dbo.Vendedores v on VC.vendedor_3 = v.codigo_ven
	where vc.flag_3 = '1' and vc.cperiodo_3 ='" + Trim(TextBox1.Text) + "'  and vc.empresa ='" + Trim(Label5.Text) + "' and MONTH(fcom_3) = ISNULL(" + gmes + ", Month(fcom_3))   and vc.vendedor_3 = '" + Trim(Label6.Text) + "' ORDER BY fcom_3", conx)
		Dim da As New DataTable
		cmd.Fill(da)
		If da.Rows.Count <> 0 Then
			DataGridView1.DataSource = da
		Else
			DataGridView1.DataSource = Nothing
		End If
		DataGridView1.DataSource = da

		Dim p As Integer
		Dim po As Double
		p = DataGridView1.Rows.Count

		po = 0
		For i = 0 To p - 1
			po = po + DataGridView1.Rows(i).Cells(5).Value
		Next

		TextBox4.Text = po

		DataGridView1.Columns(0).Width = 160
		DataGridView1.Columns(1).Width = 160
		DataGridView1.Columns(3).Width = 75
		DataGridView1.Columns(5).Width = 85
		DataGridView1.Columns(6).Width = 85
		DataGridView1.Columns(7).Width = 85
		DataGridView1.Columns(8).Width = 85
		DataGridView1.Columns(9).Width = 85
		DataGridView1.Columns(10).Width = 85
		DataGridView1.ReadOnly = True
	End Sub

	Private Sub Rpt_Ventas_ven_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		ComboBox1.SelectedIndex = DateTime.Now.ToString("MM") - 1
	End Sub

	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
		Dim expor As New Exportar
		expor.llenarExcel(DataGridView1)

	End Sub
End Class