Imports System.Data.SqlClient
Public Class Gand_Udp
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

    Dim da As New DataTable
    Sub mostrar_informacion()
        abrir()
        da.Clear()
        DataGridView1.DataSource = Nothing
        Dim cmd As New SqlDataAdapter("SELECT DISTINCT SUBSTRING(ncom_3,1,9) + '-'+ SUBSTRING(ncom_3,10,1) AS OD,ncom_3,program_3 AS PM,c.nomb_10 AS RUC,isnull((SELECT 
CASE WHEN (select top 1 EtaRut from Ruta_Udp WHERE CiRut ='1' and OdRut=ncom_3   ORDER BY EtaRut desc ) ='01' then 'CREACION MOLDE'
 WHEN (select top 1 EtaRut from Ruta_Udp WHERE CiRut ='1' and OdRut=ncom_3   ORDER BY EtaRut desc ) ='02' then 'CORTE' 
WHEN (select top 1 EtaRut from Ruta_Udp WHERE CiRut ='1' and OdRut=ncom_3   ORDER BY EtaRut desc ) ='03' then 'APLICACIONES' 
WHEN (select top 1 EtaRut from Ruta_Udp WHERE CiRut ='1' and OdRut=ncom_3   ORDER BY EtaRut desc ) ='04' then 'CONFECCION' 
WHEN (select top 1 EtaRut from Ruta_Udp WHERE CiRut ='1' and OdRut=ncom_3   ORDER BY EtaRut desc ) ='05' then 'LAVADO' 
WHEN (select top 1 EtaRut from Ruta_Udp WHERE CiRut ='1' and OdRut=ncom_3   ORDER BY EtaRut desc ) ='06' then 'ACABADOS' 
WHEN (select top 1 EtaRut from Ruta_Udp WHERE CiRut ='1' and OdRut=ncom_3   ORDER BY EtaRut desc ) ='07' then 'ENTREGA COMERCIAL' ELSE 'ANALISIS PRENDA'
END),'FALTA INFORMACION') AS ESTADO, t.dele AS MARCA,descri_3 AS DESCRIPCION,alterno_3 as LAVADO_COM,estilo_3 as ESTILO, cants_3 AS SOLICITADO, 
cantp_3 AS PROYECCION,broker_3 AS COMERCIAL,telaprinc_3 AS TELA,lavadoten_3 AS LAVADO, otros1_3 AS COLOR, otros2_3 AS EFECTOS, otros3_3 AS DIFICULTAD,
fcom_3 AS F_INICIO,ffinprod_3 AS F_FIN
FROM custom_vianny.DBO.qag0300  q left join
custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 =c.fich_10 
left join custom_vianny.dbo.Tab0100 t on t.CTab= 'MARCLI' and  t.cele =q.marcacli_3
where ncom_3 like '03%' AND q.ccia ='01'


", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(4).Width = 135
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = 250
        DataGridView1.Columns(7).Width = 80
        DataGridView1.Columns(8).Width = 70
        DataGridView1.Columns(9).Width = 70
        DataGridView1.Columns(10).Width = 80
        DataGridView1.Columns(11).Width = 90
        DataGridView1.Columns(12).Width = 250
        DataGridView1.Columns(0).Frozen = True
        DataGridView1.Columns(1).Frozen = True
        DataGridView1.Columns(2).Frozen = True
        DataGridView1.Columns(3).Frozen = True
        DataGridView1.Columns(4).Frozen = True
        DataGridView1.Columns(5).Frozen = True
    End Sub


    Private Sub Gand_Udp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mostrar_informacion()
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 0 Then
            Detalle_Status.TextBox1.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value), 1, 9) & Microsoft.VisualBasic.Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value), 11, 1)
            Detalle_Status.ShowDialog()
        End If
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "ESTADO" & " = '" & ComboBox1.SelectedValue & "%'"

            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 60
                DataGridView1.Columns(1).Width = 190
                DataGridView1.Columns(0).ReadOnly = True
                DataGridView1.Columns(1).ReadOnly = True
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
    Dim dv As New DataView

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Text = ComboBox1.Text
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        mostrar_informacion()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class