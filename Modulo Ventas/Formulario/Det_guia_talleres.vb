Imports System.Data.SqlClient
Public Class Det_guia_talleres
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
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
    Dim da1 As New DataTable
    Private Sub Det_guia_talleres_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = Nothing
        da1.Clear()
        DateTimePicker1.Value = DateTimePicker2.Value.AddDays(-30)
        abrir()
        Dim cmd55 As New SqlDataAdapter("select sfactu_3 as SERIE, nfactu_3 AS CORRELATIVO,fich_3 AS RUC,nombfich_3 AS PROVEEDOR,fcom_3 AS FECHA,CASE WHEN  fase_3 = '0700' THEN 'APLICACIONES Y PIEZAS' WHEN fase_3 ='0300' THEN 'CORTE' WHEN fase_3='0400' THEN 'CONFECCION' WHEN fase_3 ='0600' THEN 'APLICACIONES' WHEN fase_3 ='0500' THEN 'ACABADO' END AS PROCESO  from custom_vianny.dbo.Fag0400 a where A.CIA_3 = '" + TextBox2.Text + "'  AND A.fcom_3 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.fcom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' and fase_3 ='" + Label3.Text + "'", conx)

        cmd55.Fill(da1)
        If da1.Rows.Count > 0 Then
            DataGridView1.DataSource = da1

            DataGridView1.Columns(3).Width = 360
            DataGridView1.Columns(5).Width = 230
        Else
            DataGridView1.DataSource = Nothing
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = Nothing
        da1.Clear()
        abrir()
        Dim cmd55 As New SqlDataAdapter("select sfactu_3 as SERIE, nfactu_3 AS CORRELATIVO,fich_3 AS RUC,nombfich_3 AS PROVEEDOR,fcom_3 AS FECHA,CASE WHEN  fase_3 = '0700' THEN 'APLICACIONES Y PIEZAS' WHEN fase_3 ='0300' THEN 'CORTE' WHEN fase_3='0400' THEN 'CONFECCION' WHEN fase_3 ='0600' THEN 'APLICACIONES' WHEN fase_3 ='0500' THEN 'ACABADO' END AS PROCESO  from custom_vianny.dbo.Fag0400 a where A.CIA_3 = '" + TextBox2.Text + "'  AND A.fcom_3 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.fcom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' and fase_3 ='" + Label3.Text + "'", conx)

        cmd55.Fill(da1)
        If da1.Rows.Count > 0 Then
            DataGridView1.DataSource = da1

            DataGridView1.Columns(3).Width = 360
            DataGridView1.Columns(5).Width = 205
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub
    Private Sub buscar0()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da1.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PROVEEDOR" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(3).Width = 360
                DataGridView1.Columns(5).Width = 205

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar0()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        guia_talleres.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
        guia_talleres.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        guia_talleres.TextBox2.Focus()
        SendKeys.Send("{ENTER}")
        Me.Close()
    End Sub
End Class