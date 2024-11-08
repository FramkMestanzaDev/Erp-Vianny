Imports System.Data.SqlClient
Public Class Cierre_Proc_Od
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Dim da, da1, da2 As New DataTable
    Sub INFORMACION_CERRADA()
        abrir()
        da2.Clear()
        DataGridView1.DataSource = Nothing
        Dim cmd As New SqlDataAdapter("SELECT CONVERT(varchar,fcom_3,103) AS F_INICIO,program_3 AS PM, LEFT(ncom_3, 9) AS OD,SUBSTRING( ncom_3,10,1) AS VERSION,   
 c.nomb_10 AS RUC,descri_3 AS DESCRIPCION,telaprinc_3 AS TELA, cants_3 AS SOLICITADO,q.analista_3 as ANALISTA,modelista_3 as MODELISTA,estilo_3 as 'ESTILO CLIENTE', case when merchan_3 = '0033' then 'JRAMIREZ' WHEN merchan_3 = '0003' then 'GBALVIN' WHEN merchan_3 = '0032' then 'VVARGAS' WHEN merchan_3 = '0023' then 'DBRAVO' ELSE 'VGRAUS' END AS COMERCIAL,
otros3_3 AS DIFICULTAD,CONVERT(varchar,FCome1_3,23) AS F_FIN, ISNULL(EstadoRuta.Estado, 'FALTA INFORMACION') AS ESTADO   -- Usamos OUTER APPLY para evitar repetir la subconsulta,
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
where ncom_3 like '03%' AND C.ccia ='01' and finped_3 = 1 AND flag_3 ='1' ORDER BY ncom_3 desc", conx)
        cmd.Fill(da2)

        If da2.Rows.Count > 0 Then
            DataGridView1.DataSource = da2
            DataGridView1.Columns(3).Width = 73
            DataGridView1.Columns(4).Width = 77
            DataGridView1.Columns(5).Width = 77
            DataGridView1.Columns(6).Width = 28
            DataGridView1.Columns(7).Width = 28
            DataGridView1.Columns(8).Width = 150
            DataGridView1.Columns(9).Width = 150
            DataGridView1.Columns(10).Width = 28
            DataGridView1.Columns(11).Width = 68
            DataGridView1.Columns(12).Width = 68
            DataGridView1.Columns(13).Width = 87
            DataGridView1.Columns(14).Width = 78
            DataGridView1.Columns(16).Width = 65
            DataGridView1.Columns(15).Visible = False
            'DataGridView1.Columns(0).Frozen = True
            'DataGridView1.Columns(1).Frozen = True
            'DataGridView1.Columns(2).Frozen = True
            'DataGridView1.Columns(3).Frozen = True
            'DataGridView1.Columns(4).Frozen = True
            'DataGridView1.Columns(5).Frozen = True
            'DataGridView1.Columns(6).Frozen = True
            MsgBox("Resultado de : Lista de Od Cerradas")
        End If

    End Sub
    Sub mostrar_informacion()
        abrir()
        da.Clear()
        DataGridView1.DataSource = Nothing
        Dim cmd As New SqlDataAdapter("SELECT CONVERT(varchar,fcom_3,103) AS F_INICIO,program_3 AS PM, LEFT(ncom_3, 9) AS OD,SUBSTRING( ncom_3,10,1) AS VERSION,   
 c.nomb_10 AS RUC,descri_3 AS DESCRIPCION,telaprinc_3 AS TELA, cants_3 AS SOLICITADO,q.analista_3 as ANALISTA,modelista_3 as MODELISTA,estilo_3 as 'ESTILO CLIENTE', case when merchan_3 = '0033' then 'JRAMIREZ' WHEN merchan_3 = '0003' then 'GBALVIN' WHEN merchan_3 = '0032' then 'VVARGAS' WHEN merchan_3 = '0023' then 'DBRAVO' ELSE 'VGRAUS' END AS COMERCIAL,
otros3_3 AS DIFICULTAD,CONVERT(varchar,FCome1_3,23) AS F_FIN, ISNULL(EstadoRuta.Estado, 'FALTA INFORMACION') AS ESTADO   -- Usamos OUTER APPLY para evitar repetir la subconsulta,
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
where ncom_3 like '03%' AND C.ccia ='01' and finped_3 = 0 AND flag_3 ='1' ORDER BY ncom_3 desc", conx)
        cmd.Fill(da)

        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(3).Width = 73
            DataGridView1.Columns(4).Width = 77
            DataGridView1.Columns(5).Width = 77
            DataGridView1.Columns(6).Width = 28
            DataGridView1.Columns(7).Width = 28
            DataGridView1.Columns(8).Width = 150
            DataGridView1.Columns(9).Width = 150
            DataGridView1.Columns(10).Width = 28
            DataGridView1.Columns(11).Width = 68
            DataGridView1.Columns(12).Width = 68
            DataGridView1.Columns(13).Width = 87
            DataGridView1.Columns(14).Width = 78
            DataGridView1.Columns(16).Width = 65
            DataGridView1.Columns(15).Visible = False
            'DataGridView1.Columns(0).Frozen = True
            'DataGridView1.Columns(1).Frozen = True
            'DataGridView1.Columns(2).Frozen = True
            'DataGridView1.Columns(3).Frozen = True
            'DataGridView1.Columns(4).Frozen = True
            'DataGridView1.Columns(5).Frozen = True
            'DataGridView1.Columns(6).Frozen = True
            MsgBox("Resultado de : Lista de Od Pendientes")
        End If

    End Sub
    Sub Rpt_od()
        abrir()
        da1.Clear()
        DataGridView2.DataSource = Nothing
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
where ncom_3 like '03%' AND C.ccia ='01' and finped_3 = 0 AND flag_3 ='1' ORDER BY ncom_3 desc", conx)
        cmd.Fill(da1)
        DataGridView2.DataSource = da1

    End Sub

    Private Sub Cierre_Proc_Od_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True
        'mostrar_informacion()
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 1 Then
            Detalle_Cie_pro.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value) & Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
            Detalle_Cie_pro.ShowDialog()
        End If
        If e.ColumnIndex = 2 Then
            Od_Udp.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
            Od_Udp.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
            Od_Udp.ShowDialog()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim p As Integer
        p = DataGridView1.Rows.Count
        If p > 0 Then
            For i = 0 To p - 1
                If DataGridView1.Rows(i).Cells(0).Value = True Then
                    Dim cmd20 As New SqlCommand("update custom_vianny.dbo.qag0300 set  finped_3=1 where ncom_3 =@ncom_3 and ccia ='01'", conx)
                    cmd20.Parameters.AddWithValue("@ncom_3", Trim(DataGridView1.Rows(i).Cells(5).Value) & Trim(DataGridView1.Rows(i).Cells(6).Value))
                    cmd20.ExecuteNonQuery()
                End If
            Next
            MsgBox("SE CERRO LA OD")
            mostrar_informacion()
        Else
            MsgBox("NO HAY INFORMACION PARA CERRAR")

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Trim(ComboBox1.Text) = "" Then
            MsgBox("NO SELECCIONO NINGUN PROCESO PARA BUSCAR")
        Else

            Dim campo As String = "ESTADO"

            ' Filtrar los datos en el DataTable usando la expresión de filtro
            Dim filtro As String = $"{campo} LIKE '%{TextBox1.Text}%'"
            Dim resultados() As DataRow = da.Select(filtro)

            ' Mostrar los resultados en el DataGridView
            Dim dtResultados As DataTable = da.Clone()
            For Each resultado As DataRow In resultados
                dtResultados.ImportRow(resultado)
            Next

            DataGridView1.DataSource = dtResultados
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Text = ComboBox1.Text
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        mostrar_informacion()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscarpm()
    End Sub
    Private Sub buscarpm()
        Try
            Dim ds As New DataSet
            If RadioButton1.Checked = True Then
                ds.Tables.Add(da.Copy)
            Else
                ds.Tables.Add(da2.Copy)
            End If
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "PM" & " like '%" & TextBox2.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(3).Width = 73
                DataGridView1.Columns(4).Width = 77
                DataGridView1.Columns(5).Width = 77
                DataGridView1.Columns(6).Width = 28
                DataGridView1.Columns(7).Width = 28
                DataGridView1.Columns(8).Width = 150
                DataGridView1.Columns(9).Width = 150
                DataGridView1.Columns(10).Width = 28
                DataGridView1.Columns(11).Width = 68
                DataGridView1.Columns(12).Width = 68
                DataGridView1.Columns(13).Width = 87
                DataGridView1.Columns(14).Width = 78
                DataGridView1.Columns(16).Width = 65
                DataGridView1.Columns(15).Visible = False
                'DataGridView1.Columns(0).Frozen = True
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
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub buscarOd()
        Try
            Dim ds As New DataSet
            If RadioButton1.Checked = True Then
                ds.Tables.Add(da.Copy)
            Else
                ds.Tables.Add(da2.Copy)
            End If

            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "OD" & " like '%" & TextBox3.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(3).Width = 73
                DataGridView1.Columns(4).Width = 77
                DataGridView1.Columns(5).Width = 77
                DataGridView1.Columns(6).Width = 28
                DataGridView1.Columns(7).Width = 28
                DataGridView1.Columns(8).Width = 150
                DataGridView1.Columns(9).Width = 150
                DataGridView1.Columns(10).Width = 28
                DataGridView1.Columns(11).Width = 68
                DataGridView1.Columns(12).Width = 68
                DataGridView1.Columns(13).Width = 87
                DataGridView1.Columns(14).Width = 78
                DataGridView1.Columns(16).Width = 65
                DataGridView1.Columns(15).Visible = False
                'DataGridView1.Columns(0).Frozen = True
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
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            mostrar_informacion()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            INFORMACION_CERRADA()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim p As Integer
        p = DataGridView1.Rows.Count
        If p > 0 Then
            For i = 0 To p - 1
                If DataGridView1.Rows(i).Cells(0).Value = True Then
                    Dim cmd20 As New SqlCommand("update custom_vianny.dbo.qag0300 set  finped_3=0 where ncom_3 =@ncom_3 and ccia ='01'", conx)
                    cmd20.Parameters.AddWithValue("@ncom_3", Trim(DataGridView1.Rows(i).Cells(5).Value.ToString()) & Trim(DataGridView1.Rows(i).Cells(6).Value.ToString()))
                    cmd20.ExecuteNonQuery()
                End If
            Next
            MsgBox("SE CERRO LA OD")
            mostrar_informacion()
        Else
            MsgBox("NO HAY INFORMACION PARA CERRAR")

        End If
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick

    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        If e.ColumnIndex = 4 Or e.ColumnIndex = 5 Then

            Dim cellValue As String = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()

            Clipboard.SetText(cellValue)
            ' Opcional: Mostrar un mensaje de confirmación
            MessageBox.Show("El valor de la celda " & DataGridView1.Columns(e.ColumnIndex).HeaderText.ToString().Trim() & " ha sido copiado al portapapeles: " & cellValue, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Columns(e.ColumnIndex).Name = "ESTADO" Then
            If e.Value IsNot Nothing AndAlso Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value) >= 2 AndAlso Convert.ToString(e.Value.ToString().Trim()) <> "CORTE" Then
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Yellow
            Else
                If e.Value IsNot Nothing AndAlso Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value) >= 2 AndAlso Convert.ToString(e.Value.ToString().Trim()) = "CORTE" Then
                    DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightGreen
                End If

            End If
        End If
        'If DataGridView1.Columns(e.ColumnIndex).Name = "VERSION" Then
        '    If e.Value IsNot Nothing AndAlso Convert.ToInt32(e.Value.ToString()) >= 2 Then
        '        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightGreen
        '    End If
        'End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscarOd()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Rpt_od()
        Dim ext As New ExportaruDP
        ext.llenarExcel(DataGridView2)
    End Sub
End Class