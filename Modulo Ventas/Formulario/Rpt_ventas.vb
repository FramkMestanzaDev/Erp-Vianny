Imports System.Data.SqlClient
Public Class Rpt_ventas
    Public conx As SqlConnection
    Public comando As SqlCommand
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim GG As New fventas
        Dim hh As New vventas

        hh.gANO = Label7.Text
        hh.gccia = Label2.Text
        If ComboBox1.Text = "" Or ComboBox1.Text = "TODOS" Then
            hh.gmes = "NULL"

        Else
            Select Case ComboBox1.Text
                Case "ENERO" : hh.gmes = "01"
                Case "FEBRERO" : hh.gmes = "02"
                Case "MARZO" : hh.gmes = "03"
                Case "ABRIL" : hh.gmes = "04"
                Case "MAYO" : hh.gmes = "05"
                Case "JUNIO" : hh.gmes = "06"
                Case "JULIO" : hh.gmes = "07"
                Case "AGOSTO" : hh.gmes = "08"
                Case "SETIEMBRE" : hh.gmes = "09"
                Case "OCTUBRE" : hh.gmes = "10"
                Case "NOVIEMBRE" : hh.gmes = "11"
                Case "DICIEMBRE" : hh.gmes = "12"
            End Select

        End If
        Rpt_ventas_rubro.TextBox1.Text = hh.gANO
        Rpt_ventas_rubro.TextBox2.Text = hh.gccia
        Rpt_ventas_rubro.TextBox3.Text = hh.gmes
        Rpt_ventas_rubro.Show()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        VENTAS_RUBROS.TextBox1.Text = Label7.Text
        VENTAS_RUBROS.TextBox2.Text = Label2.Text
        VENTAS_RUBROS.Show()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim GG As New fventas
        Dim hh As New vventas
        Dim h As Double
        hh.gANO = Label7.Text
        hh.gccia = Label2.Text
        If ComboBox1.Text = "" Or ComboBox1.Text = "TODOS" Then
            hh.gmes = "NULL"

        Else
            Select Case ComboBox1.Text
                Case "ENERO" : hh.gmes = "01"
                Case "FEBRERO" : hh.gmes = "02"
                Case "MARZO" : hh.gmes = "03"
                Case "ABRIL" : hh.gmes = "04"
                Case "MAYO" : hh.gmes = "05"
                Case "JUNIO" : hh.gmes = "06"
                Case "JULIO" : hh.gmes = "07"
                Case "AGOSTO" : hh.gmes = "08"
                Case "SETIEMBRE" : hh.gmes = "09"
                Case "OCTUBRE" : hh.gmes = "10"
                Case "NOVIEMBRE" : hh.gmes = "11"
                Case "DICIEMBRE" : hh.gmes = "12"
            End Select

        End If
        Rpt_Pro_ven.TextBox1.Text = hh.gANO
        Rpt_Pro_ven.TextBox2.Text = hh.gccia
        Rpt_Pro_ven.TextBox3.Text = hh.gmes
        Rpt_Pro_ven.Show()


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
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        rpt_vendedo.TextBox1.Text = Label7.Text
        rpt_vendedo.TextBox2.Text = Label2.Text
        rpt_vendedo.Show()
    End Sub
    Dim KK As DataTable
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim JK As New fventas
        Dim LL As New vventas

        LL.gccia = Label2.Text
        LL.gANO = Label7.Text
        KK = JK.VENTA_TOTALES_VENDEDOR(LL)
        If KK.Rows.Count <> 0 Then
            DataGridView1.DataSource = KK
        End If
        Dim ext As New Exportar
        ext.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
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
        Dim cmd As New SqlDataAdapter("Set Language spanish  SELECT 
 fcom_3 as FECHA,a.fich_3 as RUC,
 V.alias_ven as VENDEDOR,
	 CASE WHEN A.TiDoc_3 = '007' THEN -1 ELSE 1 END *  B.VVta1_3a   AS VALOR_VENTA,
ISNULL( jk.nomb_14f,'') AS RUBRO_DETALLADO,
ISNULL( jk.nomb_14_rubroge,'') as RUBRO_GENERAL,
	 case when 	b.grm_3a = '' then a.sfactur_3+ '-'+a.nfactur_3 else 	b.grm_3a end AS GRM,
	 a.sfactu_3 AS SERIE_FAC,
	 a.nfactu_3 CORRELATIVO_FAC,
	 B.preun1_3a  AS PRECIO_S,
	 a.nomb_3  AS CLIENTE,
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
ELSE 'VACIO'
	END as ALMACEN,
	ISNULL( c.nomb_17,'') AS PRODCUTO,
	 a.tcam_3 AS T_CAMBIO,
	isnull(k.nomb_15k,'') AS FO_PAGO,
B.cant_3a AS CANTIDAD ,
A.fich_3
	 FROM custom_vianny.DBO.Fag0300 A LEFT JOIN custom_vianny.DBO.Fap0300 B
					ON A.CCia_3 = B.CCia_3a AND A.CPeriodo_3 = B.CPeriodo_3a AND A.Tienda_3 = B.Tienda_3a AND A.NCom_3 = B.NCom_3a	
					LEFT JOIN custom_vianny.DBO.cag1700 c ON b.ccia_3a=c.ccia and b.linea_3a = c.linea_17
left join custom_vianny.dbo.kag1500 k on  A.condp_3 = k.cond_15k and k.ccia_15k = '" + Trim(Label2.Text) + "'
left join custom_vianny.dbo.Vendedores v on a.vendedor_3 = v.codigo_ven
LEFT JOIN custom_vianny.DBO.fag1400  JK ON B.rubro_3a = JK.oper_14f and jk.ccia_14f ='" + Trim(Label2.Text) + "' 
	 WHERE cperiodo_3 = '" + Trim(Label7.Text) + "' AND ccia_3 ='" + Trim(Label2.Text) + "' AND tidoc_3 in (001,003,007)   AND flag_3 ='1' and MONTH(fcom_3)= ISNULL(" + gmes + ", Month(fcom_3)) and v.rpm_ven <> '' AND  B.rubro_3a not in (0035,0036)", conx)
        Dim da As New DataTable
        cmd.Fill(da)
        DataGridView2.DataSource = da
        Dim ext As New ExportarCom
        ext.llenarExcel(DataGridView2)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        PEDIDO.Show()
    End Sub

    Private Sub Rpt_ventas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = DateTime.Now.ToString("MM") - 1
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim GG As New fventas
        Dim hh As New vventas

        hh.gANO = Label7.Text
        hh.gccia = Label2.Text
        If ComboBox1.Text = "" Or ComboBox1.Text = "TODOS" Then
            hh.gmes = "NULL"

        Else
            Select Case ComboBox1.Text
                Case "ENERO" : hh.gmes = "01"
                Case "FEBRERO" : hh.gmes = "02"
                Case "MARZO" : hh.gmes = "03"
                Case "ABRIL" : hh.gmes = "04"
                Case "MAYO" : hh.gmes = "05"
                Case "JUNIO" : hh.gmes = "06"
                Case "JULIO" : hh.gmes = "07"
                Case "AGOSTO" : hh.gmes = "08"
                Case "SETIEMBRE" : hh.gmes = "09"
                Case "OCTUBRE" : hh.gmes = "10"
                Case "NOVIEMBRE" : hh.gmes = "11"
                Case "DICIEMBRE" : hh.gmes = "12"
            End Select

        End If
        Rpt_Ventas_rubro_vendedor.TextBox1.Text = hh.gANO
        Rpt_Ventas_rubro_vendedor.TextBox2.Text = hh.gccia
        Rpt_Ventas_rubro_vendedor.TextBox3.Text = hh.gmes
        Rpt_Ventas_rubro_vendedor.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim GG As New fventas
        Dim hh As New vventas

        hh.gANO = Label7.Text
        hh.gccia = Label2.Text
        If ComboBox1.Text = "" Or ComboBox1.Text = "TODOS" Then
            hh.gmes = "NULL"

        Else
            Select Case ComboBox1.Text
                Case "ENERO" : hh.gmes = "01"
                Case "FEBRERO" : hh.gmes = "02"
                Case "MARZO" : hh.gmes = "03"
                Case "ABRIL" : hh.gmes = "04"
                Case "MAYO" : hh.gmes = "05"
                Case "JUNIO" : hh.gmes = "06"
                Case "JULIO" : hh.gmes = "07"
                Case "AGOSTO" : hh.gmes = "08"
                Case "SETIEMBRE" : hh.gmes = "09"
                Case "OCTUBRE" : hh.gmes = "10"
                Case "NOVIEMBRE" : hh.gmes = "11"
                Case "DICIEMBRE" : hh.gmes = "12"
            End Select

        End If
        Rpt_ventas_cliente.TextBox1.Text = hh.gANO
        Rpt_ventas_cliente.TextBox2.Text = hh.gccia
        Rpt_ventas_cliente.TextBox3.Text = hh.gmes
        Rpt_ventas_cliente.Show()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Cliente_nuevos.Show()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        rOTACION_CLIENTE.TextBox1.Text = Label4.Text
        rOTACION_CLIENTE.TextBox2.Text = Label7.Text
        rOTACION_CLIENTE.Show()
    End Sub
End Class