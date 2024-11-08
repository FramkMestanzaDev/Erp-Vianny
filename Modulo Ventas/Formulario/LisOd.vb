Imports System.Data.SqlClient
Imports System.Windows.Forms
Public Class LisOd
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
        buscar_OD()
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
    Private Sub buscar_OD()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT CONVERT(varchar,fcom_3,103) AS F_INICIO,program_3 AS PM, LEFT(ncom_3, 9) AS OD,SUBSTRING( ncom_3,10,1) AS VERSION,   
 c.nomb_10 AS RUC,descri_3 AS DESCRIPCION,telaprinc_3 AS TELA, cants_3 AS SOLICITADO,q.analista_3 as ANALISTA,modelista_3 as MODELISTA, case when merchan_3 = '0033' then 'JRAMIREZ' WHEN merchan_3 = '0003' then 'GBALVIN' WHEN merchan_3 = '0032' then 'VVARGAS' WHEN merchan_3 = '0023' then 'DBRAVO' ELSE 'VGRAUS' END AS COMERCIAL,
estilo_3 as 'ESTILO CLIENTE',otros3_3 AS DIFICULTAD,CONVERT(varchar,FCome1_3,23) AS F_FIN, ISNULL(EstadoRuta.Estado, 'FALTA INFORMACION') AS ESTADO   -- Usamos OUTER APPLY para evitar repetir la subconsulta,
FROM custom_vianny.DBO.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and (q.fich_3 = c.fich_10 or q.fich_3 =c.ruc_10)
OUTER APPLY 
    (SELECT 
        CASE 
            WHEN EtaRut = '01' THEN 'CREACION MOLDE'
            WHEN EtaRut = '02' THEN 'CORTE'
            WHEN EtaRut = '03' THEN 'APLICACIONES'
            WHEN EtaRut = '04' THEN 'CONFECCION'
            WHEN EtaRut = '05' THEN 'LAVADO'
            WHEN EtaRut = '06' THEN 'ACABADOS'
            WHEN EtaRut = '07' OR EtaRut = '08' THEN 'ENTREGA COMERCIAL'
            ELSE 'ANALISIS PRENDA'
        END AS Estado
     FROM 
        (SELECT TOP 1 EtaRut 
         FROM Ruta_Udp 
         WHERE CiRut = '1' AND OdRut = ncom_3 
         ORDER BY EtaRut DESC
        ) AS UltimaRuta
    ) AS EstadoRuta
where ncom_3 like '03%' AND C.ccia ='01' and finped_3 = 0 AND flag_3 ='1' AND merchan_3 ='" + Trim(Label2.Text) + "' ORDER BY ncom_3 desc", conx)
        cmd.Fill(da)

        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            'DataGridView1.Columns(1).Width = 70
            'DataGridView1.Columns(2).Width = 70
            'DataGridView1.Columns(3).Width = 75
            'DataGridView1.Columns(4).Width = 75
            'DataGridView1.Columns(5).Width = 60
            'DataGridView1.Columns(6).Width = 280
            'DataGridView1.Columns(7).Visible = False
            'DataGridView1.Columns(8).Width = 150
            'DataGridView1.Columns(9).Width = 80
            'DataGridView1.Columns(10).Width = 80
            'DataGridView1.Columns(11).Width = 90
            'DataGridView1.Columns(12).Width = 75
            'DataGridView1.Columns(13).Width = 200
            'DataGridView1.Columns(14).Width = 100
            'DataGridView1.Columns(1).Frozen = True
            'DataGridView1.Columns(2).Frozen = True
            'DataGridView1.Columns(3).Frozen = True
            'DataGridView1.Columns(4).Frozen = True
            'DataGridView1.Columns(5).Frozen = True
            'DataGridView1.Columns(6).Frozen = True
            MsgBox("Resultado de : Lista de Od Pendientes")
        End If
    End Sub
    Private Sub buscar_OD_cerrado()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT CONVERT(varchar,fcom_3,103) AS F_INICIO,program_3 AS PM, LEFT(ncom_3, 9) AS OD,SUBSTRING( ncom_3,10,1) AS VERSION,   
 c.nomb_10 AS RUC,descri_3 AS DESCRIPCION,telaprinc_3 AS TELA, cants_3 AS SOLICITADO,q.analista_3 as ANALISTA,modelista_3 as MODELISTA, case when merchan_3 = '0033' then 'JRAMIREZ' WHEN merchan_3 = '0003' then 'GBALVIN' WHEN merchan_3 = '0032' then 'VVARGAS' WHEN merchan_3 = '0023' then 'DBRAVO' ELSE 'VGRAUS' END AS COMERCIAL,
estilo_3 as 'ESTILO CLIENTE',otros3_3 AS DIFICULTAD,CONVERT(varchar,FCome1_3,23) AS F_FIN, ISNULL(EstadoRuta.Estado, 'FALTA INFORMACION') AS ESTADO   -- Usamos OUTER APPLY para evitar repetir la subconsulta,
FROM custom_vianny.DBO.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and (q.fich_3 = c.fich_10 or q.fich_3 =c.ruc_10)
OUTER APPLY 
    (SELECT 
        CASE 
            WHEN EtaRut = '01' THEN 'CREACION MOLDE'
            WHEN EtaRut = '02' THEN 'CORTE'
            WHEN EtaRut = '03' THEN 'APLICACIONES'
            WHEN EtaRut = '04' THEN 'CONFECCION'
            WHEN EtaRut = '05' THEN 'LAVADO'
            WHEN EtaRut = '06' THEN 'ACABADOS'
            WHEN EtaRut = '07' OR EtaRut = '08' THEN 'ENTREGA COMERCIAL'
            ELSE 'ANALISIS PRENDA'
        END AS Estado
     FROM 
        (SELECT TOP 1 EtaRut 
         FROM Ruta_Udp 
         WHERE CiRut = '1' AND OdRut = ncom_3 
         ORDER BY EtaRut DESC
        ) AS UltimaRuta
    ) AS EstadoRuta
where ncom_3 like '03%' AND C.ccia ='01' and finped_3 = 1 AND flag_3 ='1' AND merchan_3 ='" + Trim(Label2.Text) + "' ORDER BY ncom_3 desc", conx)
        cmd.Fill(da)

        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            'DataGridView1.Columns(1).Width = 70
            'DataGridView1.Columns(2).Width = 70
            'DataGridView1.Columns(3).Width = 75
            'DataGridView1.Columns(4).Width = 75
            'DataGridView1.Columns(5).Width = 60
            'DataGridView1.Columns(6).Width = 280
            'DataGridView1.Columns(7).Visible = False
            'DataGridView1.Columns(8).Width = 150
            'DataGridView1.Columns(9).Width = 80
            'DataGridView1.Columns(10).Width = 80
            'DataGridView1.Columns(11).Width = 90
            'DataGridView1.Columns(12).Width = 75
            'DataGridView1.Columns(13).Width = 200
            'DataGridView1.Columns(14).Width = 100
            'DataGridView1.Columns(1).Frozen = True
            'DataGridView1.Columns(2).Frozen = True
            'DataGridView1.Columns(3).Frozen = True
            'DataGridView1.Columns(4).Frozen = True
            'DataGridView1.Columns(5).Frozen = True
            'DataGridView1.Columns(6).Frozen = True
            MsgBox("Resultado de : Lista de Od Cerradas")
        End If
    End Sub
    Private Sub LisOd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True
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

            dv.RowFilter = "PM" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                'DataGridView1.Columns(1).Width = 70
                'DataGridView1.Columns(2).Width = 70
                'DataGridView1.Columns(3).Width = 75
                'DataGridView1.Columns(4).Width = 75
                'DataGridView1.Columns(5).Width = 60
                'DataGridView1.Columns(6).Width = 280
                'DataGridView1.Columns(7).Visible = False
                'DataGridView1.Columns(8).Width = 150
                'DataGridView1.Columns(9).Width = 80
                'DataGridView1.Columns(10).Width = 80
                'DataGridView1.Columns(11).Width = 90
                'DataGridView1.Columns(12).Width = 75
                'DataGridView1.Columns(13).Width = 200
                'DataGridView1.Columns(14).Width = 100
                'DataGridView1.Columns(1).Frozen = True
                'DataGridView1.Columns(2).Frozen = True
                'DataGridView1.Columns(3).Frozen = True
                'DataGridView1.Columns(4).Frozen = True
                'DataGridView1.Columns(5).Frozen = True
                'DataGridView1.Columns(6).Frozen = True
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
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
                'DataGridView1.Columns(1).Width = 70
                'DataGridView1.Columns(2).Width = 70
                'DataGridView1.Columns(3).Width = 75
                'DataGridView1.Columns(4).Width = 75
                'DataGridView1.Columns(5).Width = 60
                'DataGridView1.Columns(6).Width = 280
                'DataGridView1.Columns(7).Visible = False
                'DataGridView1.Columns(8).Width = 150
                'DataGridView1.Columns(9).Width = 80
                'DataGridView1.Columns(10).Width = 80
                'DataGridView1.Columns(11).Width = 90
                'DataGridView1.Columns(12).Width = 75
                'DataGridView1.Columns(13).Width = 200
                'DataGridView1.Columns(14).Width = 100
                'DataGridView1.Columns(1).Frozen = True
                'DataGridView1.Columns(2).Frozen = True
                'DataGridView1.Columns(3).Frozen = True
                'DataGridView1.Columns(4).Frozen = True
                'DataGridView1.Columns(5).Frozen = True
                'DataGridView1.Columns(6).Frozen = True
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
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

            dv.RowFilter = "OD" & " like '%" & TextBox4.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                'DataGridView1.Columns(1).Width = 70
                'DataGridView1.Columns(2).Width = 70
                'DataGridView1.Columns(3).Width = 75
                'DataGridView1.Columns(4).Width = 75
                'DataGridView1.Columns(5).Width = 60
                'DataGridView1.Columns(6).Width = 280
                'DataGridView1.Columns(7).Visible = False
                'DataGridView1.Columns(8).Width = 150
                'DataGridView1.Columns(9).Width = 80
                'DataGridView1.Columns(10).Width = 80
                'DataGridView1.Columns(11).Width = 90
                'DataGridView1.Columns(12).Width = 75
                'DataGridView1.Columns(13).Width = 200
                'DataGridView1.Columns(14).Width = 100
                'DataGridView1.Columns(1).Frozen = True
                'DataGridView1.Columns(2).Frozen = True
                'DataGridView1.Columns(3).Frozen = True
                'DataGridView1.Columns(4).Frozen = True
                'DataGridView1.Columns(5).Frozen = True
                'DataGridView1.Columns(6).Frozen = True
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'If e.ColumnIndex = 0 Then

        '    ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
        '    Dim num1, num2 As Integer
        '    num1 = e.RowIndex.ToString
        '    num2 = e.ColumnIndex.ToString

        '    OD.TextBox15.Text = Trim(TextBox1.Text)
        '    If Label3.Text = "01" Then
        '        OD.TextBox7.Text = "VIANNY"
        '    Else
        '        OD.TextBox7.Text = "GRAUS"
        '    End If
        '    OD.Label23.Text = Label3.Text
        '    OD.TextBox22.Text = Trim(DataGridView1.Rows(num1).Cells(4).Value)
        '    OD.TextBox23.Text = Trim(DataGridView1.Rows(num1).Cells(5).Value)

        '    OD.Label48.Text = 1
        '    OD.ShowDialog()
        '    OD.TextBox23.Focus()
        '    SendKeys.Send("{ENTER}")
        'End If
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscar1()
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

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            buscar_OD_cerrado()
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            buscar_OD()
        End If

    End Sub
End Class