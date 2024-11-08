Imports System.Data.SqlClient
Public Class ListaHojaCotizacion
    Dim da1 As New DataTable
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub mostar_todos()
        abrir()
        da1.Clear()
        DataGridView1.DataSource = ""
        Dim cmd As New SqlDataAdapter("select substring(ncom_3,1,9) as OD,substring(ncom_3,10,1) AS VERSION,program_3 AS PO, descri_3 AS PRODUCTO,telaprinc_3 AS TELA,estilo_3 as ESTILO,c.nomb_10 as CLIENTE,case when mone_3 = 0 then 'SOLES' ELSE 'DOLARES' END as MONEDA,isnull( h.id_cotizacion,'') as COTIZACION,tela1_3,q.modelista_3 as MODELISTA,q.analista_3 AS ANALISTA,
case when ISNULL(h.estado,0) = 0 then 'PENDIENTE' when ISNULL(h.estado,0) = 1 then 'TRABAJANDO' when ISNULL(h.estado,0) = 2 then 'CERRADO' when ISNULL(h.estado,0) = 3 then 'TRABAJANDO' when ISNULL(h.estado,0) = 4 then 'PENDIENTE' when ISNULL(h.estado,0) = 5 then 'TRABAJANDO' END as ESTADO
from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 and c.ccia ='01'
LEFT JOIN HojaCotizacion H ON Q.ncom_3 = H.od+ H.od_version
where ncom_3 LIKE '03%' and fcom_3 >='20240928' and  h.AprHco = 0 ORDER BY  ncom_3 desc", conx)
        cmd.Fill(da1)
        DataGridView1.DataSource = da1
        DataGridView1.Columns(12).Visible = False
        DataGridView1.Columns(0).ReadOnly = False
        DataGridView1.Columns(3).Width = 75
        DataGridView1.Columns(4).Width = 55
        DataGridView1.Columns(5).Width = 78
        DataGridView1.Columns(6).Width = 197
        DataGridView1.Columns(7).Width = 197
        DataGridView1.Columns(8).Width = 97
        DataGridView1.Columns(9).Width = 197
        DataGridView1.Columns(10).Width = 68
        DataGridView1.Columns(11).Width = 78
        DataGridView1.Columns(13).Width = 78
        DataGridView1.Columns(14).Width = 78
        DataGridView1.Columns(15).Width = 84
        DataGridView1.Columns(3).ReadOnly = True
        DataGridView1.Columns(4).ReadOnly = True
        DataGridView1.Columns(5).ReadOnly = True
        DataGridView1.Columns(6).ReadOnly = True
        DataGridView1.Columns(7).ReadOnly = True
        DataGridView1.Columns(8).ReadOnly = True
        DataGridView1.Columns(9).ReadOnly = True
        DataGridView1.Columns(10).ReadOnly = True
        DataGridView1.Columns(11).ReadOnly = True
        DataGridView1.Columns(12).ReadOnly = True
        DataGridView1.Columns(13).ReadOnly = True
        DataGridView1.Columns(14).ReadOnly = True
        DataGridView1.Columns(15).ReadOnly = True

    End Sub
    Private Sub mostar_todos_cerrados()
        abrir()
        da1.Clear()
        DataGridView1.DataSource = ""
        Dim cmd As New SqlDataAdapter("select substring(ncom_3,1,9) as OD,substring(ncom_3,10,1) AS VERSION,program_3 AS PO, descri_3 AS PRODUCTO,telaprinc_3 AS TELA,estilo_3 as ESTILO,c.nomb_10 as CLIENTE,case when mone_3 = 0 then 'SOLES' ELSE 'DOLARES' END as MONEDA,isnull( h.id_cotizacion,'') as COTIZACION,tela1_3,q.modelista_3 as MODELISTA,q.analista_3 AS ANALISTA,
case when ISNULL(h.estado,0) = 0 then 'PENDIENTE' when ISNULL(h.estado,0) = 1 then 'TRABAJANDO' when ISNULL(h.estado,0) = 2 then 'CERRADO' when ISNULL(h.estado,0) = 3 then 'TRABAJANDO' when ISNULL(h.estado,0) = 4 then 'PENDIENTE' when ISNULL(h.estado,0) = 5 then 'TRABAJANDO' END as ESTADO
from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 and c.ccia ='01'
LEFT JOIN HojaCotizacion H ON Q.ncom_3 = H.od+ H.od_version
where ncom_3 LIKE '03%' and fcom_3 >='20240928' and  h.AprHco =1 ORDER BY  ncom_3 desc", conx)
        cmd.Fill(da1)
        DataGridView1.DataSource = da1
        DataGridView1.Columns(12).Visible = False
        DataGridView1.Columns(0).ReadOnly = False
        DataGridView1.Columns(3).Width = 75
        DataGridView1.Columns(4).Width = 55
        DataGridView1.Columns(5).Width = 78
        DataGridView1.Columns(6).Width = 197
        DataGridView1.Columns(7).Width = 197
        DataGridView1.Columns(8).Width = 97
        DataGridView1.Columns(9).Width = 197
        DataGridView1.Columns(10).Width = 68
        DataGridView1.Columns(11).Width = 78
        DataGridView1.Columns(13).Width = 78
        DataGridView1.Columns(14).Width = 78
        DataGridView1.Columns(15).Width = 84
        DataGridView1.Columns(3).ReadOnly = True
        DataGridView1.Columns(4).ReadOnly = True
        DataGridView1.Columns(5).ReadOnly = True
        DataGridView1.Columns(6).ReadOnly = True
        DataGridView1.Columns(7).ReadOnly = True
        DataGridView1.Columns(8).ReadOnly = True
        DataGridView1.Columns(9).ReadOnly = True
        DataGridView1.Columns(10).ReadOnly = True
        DataGridView1.Columns(11).ReadOnly = True
        DataGridView1.Columns(12).ReadOnly = True
        DataGridView1.Columns(13).ReadOnly = True
        DataGridView1.Columns(14).ReadOnly = True
        DataGridView1.Columns(15).ReadOnly = True

    End Sub
    Private Sub ListaHojaCotizacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        mostar_todos()
        RadioButton1.Checked = True
        Button1.Visible = True
        Button3.Visible = False
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "IMP" Then
            If DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString().Trim().Length() = 0 Then
                MsgBox("ESTA OD NO ESTA ASIGANADA A UNA COTIZACION NO SE PUEDE IMPRIMIR")
            Else
                RptFormCotizacion.TextBox2.Text = "2"
                RptFormCotizacion.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(11).Value.ToString().Trim()
                RptFormCotizacion.ShowDialog()
            End If

        Else
            If DataGridView1.Columns(e.ColumnIndex).HeaderText = "HC" Then
                If DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString().Trim().Length() = 0 Then
                    MsgBox("ESTA OD NO ESTA ASIGANADA A UNA COTIZACION NO SE PUEDE INGRESAR")
                Else
                    Dim cot As New CotizacionUdp
                    cot.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
                    cot.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value)
                    cot.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
                    cot.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(9).Value)
                    cot.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(8).Value)
                    cot.TextBox6.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value)
                    cot.TextBox8.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value)
                    cot.TextBox9.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
                    cot.TextBox14.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value)
                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(9).Value) = "0" Then
                        cot.TextBox11.Text = "SOLES"
                    Else
                        cot.TextBox11.Text = "DOLARES"
                    End If

                    cot.TextBox10.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value)
                    cot.Label11.Text = "1"
                    cot.GroupBox7.Enabled = False
                    cot.Label12.Text = "4"
                    cot.TextBox10.Focus()
                    SendKeys.Send("{ENTER}")
                    cot.ShowDialog()
                End If

            End If
        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Columns(e.ColumnIndex).Name = "ESTADO" Then
            ' Verificar si el valor de la celda es "2"

            If DataGridView1.Rows(e.RowIndex).Cells(15).Value.ToString().Trim() = "PENDIENTE" Then
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightGreen
            Else
                If DataGridView1.Rows(e.RowIndex).Cells(15).Value.ToString().Trim() = "CERRADO" Then
                    DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Black
                    DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White

                End If

            End If



        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da1.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "OD" & " like '%" & TextBox1.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(0).ReadOnly = False
                DataGridView1.Columns(3).Width = 75
                DataGridView1.Columns(4).Width = 55
                DataGridView1.Columns(5).Width = 78
                DataGridView1.Columns(6).Width = 197
                DataGridView1.Columns(7).Width = 197
                DataGridView1.Columns(8).Width = 97
                DataGridView1.Columns(9).Width = 197
                DataGridView1.Columns(10).Width = 68
                DataGridView1.Columns(11).Width = 78
                DataGridView1.Columns(13).Width = 78
                DataGridView1.Columns(14).Width = 78
                DataGridView1.Columns(15).Width = 84
                DataGridView1.Columns(3).ReadOnly = True
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
                DataGridView1.Columns(7).ReadOnly = True
                DataGridView1.Columns(8).ReadOnly = True
                DataGridView1.Columns(9).ReadOnly = True
                DataGridView1.Columns(10).ReadOnly = True
                DataGridView1.Columns(11).ReadOnly = True
                DataGridView1.Columns(12).ReadOnly = True
                DataGridView1.Columns(13).ReadOnly = True
                DataGridView1.Columns(14).ReadOnly = True
                DataGridView1.Columns(15).ReadOnly = True
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
    Private Sub buscarPo()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da1.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "PO" & " like '%" & TextBox2.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(0).ReadOnly = False
                DataGridView1.Columns(3).Width = 75
                DataGridView1.Columns(4).Width = 55
                DataGridView1.Columns(5).Width = 78
                DataGridView1.Columns(6).Width = 197
                DataGridView1.Columns(7).Width = 197
                DataGridView1.Columns(8).Width = 97
                DataGridView1.Columns(9).Width = 197
                DataGridView1.Columns(10).Width = 68
                DataGridView1.Columns(11).Width = 78
                DataGridView1.Columns(13).Width = 78
                DataGridView1.Columns(14).Width = 78
                DataGridView1.Columns(15).Width = 84
                DataGridView1.Columns(3).ReadOnly = True
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
                DataGridView1.Columns(7).ReadOnly = True
                DataGridView1.Columns(8).ReadOnly = True
                DataGridView1.Columns(9).ReadOnly = True
                DataGridView1.Columns(10).ReadOnly = True
                DataGridView1.Columns(11).ReadOnly = True
                DataGridView1.Columns(12).ReadOnly = True
                DataGridView1.Columns(13).ReadOnly = True
                DataGridView1.Columns(14).ReadOnly = True
                DataGridView1.Columns(15).ReadOnly = True
            Else
                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscarPo()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
        Dim cmd20 As New SqlCommand("update HojaCotizacion set AprHco= '1' where id_cotizacion =@id_cotizacion ", conx)
        cmd20.Parameters.AddWithValue("@id_cotizacion", filaSeleccionada.Cells("COTIZACION").Value)
        cmd20.ExecuteNonQuery()
        MsgBox("SE APROBO CORRECTAMENTE")
        mostar_todos()
        Button1.Visible = True
        Button3.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        abrir()

        Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
        Dim cmd20 As New SqlCommand("update HojaCotizacion set AprHco='0' where id_cotizacion =@id_cotizacion ", conx)
        cmd20.Parameters.AddWithValue("@id_cotizacion", filaSeleccionada.Cells("COTIZACION").Value)
        cmd20.ExecuteNonQuery()
        MsgBox("SE DESAPROBO CORRECTAMENTE")
        mostar_todos_cerrados()
        Button1.Visible = False
        Button3.Visible = True
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        mostar_todos()
        Button1.Visible = True
        Button3.Visible = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        mostar_todos_cerrados()
        RadioButton2.Checked = True
        Button1.Visible = False
        Button3.Visible = True
    End Sub
End Class