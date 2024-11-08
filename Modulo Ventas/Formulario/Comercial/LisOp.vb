Imports System.Data.SqlClient
Public Class LisOp
    Dim da As New DataTable
    Public conx As SqlConnection
    Dim Rsr212 As SqlDataReader
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
        buscar_OP()
    End Sub
    Private Sub BUSCAR_CODIGO_VENDEDOR()
        Dim CodVen As String
        CodVen = ""
        abrir()
        Dim sql1021 As String = "Select cele FROM custom_vianny.DBO.TAB0100 A  Where ccia='" + Label3.Text + "' AND CTAB='MERCHA' and codigo ='" + Trim(TextBox1.Text) + "'"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr212 = cmd1021.ExecuteReader()
        If Rsr212.Read() = True Then
            CodVen = Rsr212(0)
        End If
        Rsr212.Close()
        Label2.Text = CodVen
    End Sub
    Private Sub buscar_OP()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT CONVERT(varchar,fcom_3,23) AS 'FECHA INICIO',CONVERT(varchar,FCome1_3,23) AS 'FECHA ENTREGA', program_3 AS PO,ncom_3 AS OP,SUBSTRING( ncom_3,1,9) AS OD,c.nomb_10 AS CLIENTE, marcacli_3
,descri_3 AS DESCRIPCION,estilo_3 as 'ESTILO CLIENTE', cants_3 AS SOLICITADO,q.otros1_3 AS COLOR,q.alterno_3 as LAVADO ,case when merchan_3 = '0001' then 'VSILVERIO' WHEN merchan_3 = '0003' then 'GBALVIN' WHEN merchan_3 = '0031' then 'JPOZZI' WHEN merchan_3 = '0033' then 'JRAMIREZ' ELSE 'VVARGAR' END AS COMERCIAL,
telaprinc_3 AS TELA,(SELECT CASE WHEN    MAX(EtaRut)= '01' THEN 'ANALISIS PRENDA' WHEN   MAX(EtaRut)= '02' THEN 'CREACION MOLDE' WHEN   MAX(EtaRut)= '03' THEN 'CORTE' WHEN   MAX(EtaRut)= '04' THEN 'APLICACIONES' 
WHEN    MAX(EtaRut)= '05' THEN 'CONFECCION' WHEN    MAX(EtaRut)= '06' THEN 'LAVADO' WHEN    MAX(EtaRut)= '07' THEN 'ACABADOS' WHEN    MAX(EtaRut)= '08' THEN 'ENTREGA COMERCIAL' else 'PENDIENTE'  END AS UBICACION FROM Ruta_Udp WHERE OdRut =Q.ncom_3 AND CiRut ='1')AS UBICACION
FROM custom_vianny.DBO.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and (q.fich_3 = c.fich_10 or q.fich_3 =c.ruc_10)
where ncom_3 like '01%' AND C.ccia ='01' and finped_3 = 0 AND flag_3 ='1' AND merchan_3 ='" + Trim(Label2.Text) + "' ORDER BY ncom_3", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        If da.Rows.Count > 0 Then
            DataGridView1.Columns(1).Width = 70
            DataGridView1.Columns(2).Width = 70
            DataGridView1.Columns(3).Width = 75
            DataGridView1.Columns(4).Width = 75
            DataGridView1.Columns(5).Width = 60
            DataGridView1.Columns(6).Width = 280
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(8).Width = 150
            DataGridView1.Columns(9).Width = 80
            DataGridView1.Columns(10).Width = 80
            DataGridView1.Columns(11).Width = 90
            DataGridView1.Columns(12).Width = 75
            DataGridView1.Columns(13).Width = 200
            DataGridView1.Columns(14).Width = 100
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            DataGridView1.Columns(4).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            DataGridView1.Columns(6).Frozen = True
        End If
    End Sub

    Private Sub LisOp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BUSCAR_CODIGO_VENDEDOR()
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "PO" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 70
                DataGridView1.Columns(2).Width = 70
                DataGridView1.Columns(3).Width = 75
                DataGridView1.Columns(4).Width = 75
                DataGridView1.Columns(5).Width = 60
                DataGridView1.Columns(6).Width = 280
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Width = 150
                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Width = 80
                DataGridView1.Columns(11).Width = 90
                DataGridView1.Columns(12).Width = 75
                DataGridView1.Columns(13).Width = 200
                DataGridView1.Columns(14).Width = 100
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(6).Frozen = True
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "CLIENTE" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 70
                DataGridView1.Columns(2).Width = 70
                DataGridView1.Columns(3).Width = 75
                DataGridView1.Columns(4).Width = 75
                DataGridView1.Columns(5).Width = 60
                DataGridView1.Columns(6).Width = 280
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Width = 150
                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Width = 80
                DataGridView1.Columns(11).Width = 90
                DataGridView1.Columns(12).Width = 75
                DataGridView1.Columns(13).Width = 200
                DataGridView1.Columns(14).Width = 100
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(6).Frozen = True
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscar1()
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "OP" & " like '%" & TextBox4.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 70
                DataGridView1.Columns(2).Width = 70
                DataGridView1.Columns(3).Width = 75
                DataGridView1.Columns(4).Width = 75
                DataGridView1.Columns(5).Width = 60
                DataGridView1.Columns(6).Width = 280
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Width = 150
                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Width = 80
                DataGridView1.Columns(11).Width = 90
                DataGridView1.Columns(12).Width = 75
                DataGridView1.Columns(13).Width = 200
                DataGridView1.Columns(14).Width = 100
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(6).Frozen = True
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        buscar2()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim exp = New ExportarSoloData
        exp.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString

            Op_Manufactura.TextBox10.Text = Trim(TextBox1.Text)
            If Label3.Text = "01" Then
                Op_Manufactura.TextBox11.Text = "VIANNY"
            Else
                Op_Manufactura.TextBox11.Text = "GRAUS"
            End If
            Op_Manufactura.Label33.Text = Label3.Text
            Op_Manufactura.TextBox1.Text = Trim(DataGridView1.Rows(num1).Cells(4).Value)


            'Op_Manufactura.Label48.Text = 1
            Op_Manufactura.ShowDialog()

            Op_Manufactura.TextBox1.Focus()
            SendKeys.Send("{ENTER}")
        End If
    End Sub
End Class