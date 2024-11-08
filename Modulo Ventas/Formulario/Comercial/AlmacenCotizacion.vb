Imports System.Data.SqlClient
Public Class AlmacenCotizacion
    Dim da1 As New DataTable
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub
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
        Dim cmd As New SqlDataAdapter("select substring(ncom_3,1,9) as OD,substring(ncom_3,10,1) AS VERSION,program_3 AS PO, descri_3 AS PRODUCTO,telaprinc_3 AS TELA,mruta1_3 AS RUTA,estilo_3 as ESTILO,c.nomb_10 as CLIENTE,case when mone_3 = 0 then 'SOLES' ELSE 'DOLARES' END as MONEDA,isnull( h.id_cotizacion,'') as COTIZACION,tela1_3,
case when ISNULL(h.estado,0) = 0 then 'PENDIENTE' when ISNULL(h.estado,0) = 1 then 'TRABAJANDO' when ISNULL(h.estado,0) = 2 then 'CERRADO' when ISNULL(h.estado,0) = 3 then 'TRABAJANDO' when ISNULL(h.estado,0) = 4 then 'PENDIENTE' when ISNULL(h.estado,0) = 5 then 'TRABAJANDO' END as ESTADO,analista_3 as ANALISTA
from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 and c.ccia ='01'
LEFT JOIN HojaCotizacion H ON Q.ncom_3 = H.od+ H.od_version
where ncom_3 LIKE '03%' and fcom_3 >='20240928'  and q.finped_3 ='0' ORDER BY  ncom_3 asc", conx)
        cmd.Fill(da1)
        DataGridView1.DataSource = da1
        DataGridView1.Columns(13).Visible = False

        DataGridView1.Columns(3).Width = 75
        DataGridView1.Columns(4).Width = 65
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = 213
        DataGridView1.Columns(7).Width = 115
        DataGridView1.Columns(8).Width = 213
        DataGridView1.Columns(11).Width = 62
        DataGridView1.Columns(12).Width = 73
        DataGridView1.Columns(14).Width = 80
        DataGridView1.Columns(15).Width = 95
    End Sub
    Private Sub mostar_todos_cerrado()
        abrir()
        da1.Clear()
        DataGridView1.DataSource = ""
        Dim cmd As New SqlDataAdapter("select substring(ncom_3,1,9) as OD,substring(ncom_3,10,1) AS VERSION,program_3 AS PO, descri_3 AS PRODUCTO,telaprinc_3 AS TELA,mruta1_3 AS RUTA,estilo_3 as ESTILO,c.nomb_10 as CLIENTE,case when mone_3 = 0 then 'SOLES' ELSE 'DOLARES' END as MONEDA,isnull( h.id_cotizacion,'') as COTIZACION,tela1_3,
case when ISNULL(h.estado,0) = 0 then 'PENDIENTE' when ISNULL(h.estado,0) = 1 then 'TRABAJANDO' when ISNULL(h.estado,0) = 2 then 'CERRADO' when ISNULL(h.estado,0) = 3 then 'TRABAJANDO' when ISNULL(h.estado,0) = 4 then 'PENDIENTE' when ISNULL(h.estado,0) = 5 then 'TRABAJANDO' END as ESTADO,analista_3 as ANALISTA
from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 and c.ccia ='01'
LEFT JOIN HojaCotizacion H ON Q.ncom_3 = H.od+ H.od_version
where ncom_3 LIKE '03%' and fcom_3 >='20240928'  and q.finped_3 ='1' ORDER BY  ncom_3 asc", conx)
        cmd.Fill(da1)
        If RadioButton2.Checked = True Then
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = da1.Rows.Count

            For I = 0 To da1.Rows.Count - 1
                ProgressBar1.Visible = True
                Label5.Visible = True
                ProgressBar1.Value = I + 1
                Label5.Text = (Convert.ToInt32(ProgressBar1.Value * 100) / ProgressBar1.Maximum) & " % "
            Next
            If ProgressBar1.Value = ProgressBar1.Maximum Then
                MsgBox("Se Termino de Cargar la Informacion")
                ProgressBar1.Visible = False
                Label5.Visible = False
            End If
        End If
        DataGridView1.DataSource = da1

        DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(3).Width = 75
        DataGridView1.Columns(4).Width = 65
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = 213
        DataGridView1.Columns(7).Width = 115
        DataGridView1.Columns(8).Width = 213
        DataGridView1.Columns(11).Width = 62
        DataGridView1.Columns(12).Width = 73
        DataGridView1.Columns(14).Width = 80
        DataGridView1.Columns(15).Width = 95


    End Sub
    Private Sub mostrar_compras()
        abrir()
        da1.Clear()
        DataGridView1.DataSource = ""
        Dim cmd As New SqlDataAdapter("select substring(ncom_3,1,9) as OD,substring(ncom_3,10,1) AS VERSION,program_3 AS PO, descri_3 AS PRODUCTO,telaprinc_3 AS TELA,mruta1_3 AS RUTA,estilo_3 as ESTILO,c.nomb_10 as CLIENTE,case when mone_3 = 0 then 'SOLES' ELSE 'DOLARES' END as MONEDA,isnull( h.id_cotizacion,'') as COTIZACION,tela1_3,case when ISNULL(h.estado,0) = 0 then 'PENDIENTE' when ISNULL(h.estado,0) = 1 then 'TRABAJANDO' when ISNULL(h.estado,0) = 2 then 'CERRADO' when ISNULL(h.estado,0) = 3 then 'TRABAJANDO' when ISNULL(h.estado,0) = 4 then 'PENDIENTE' when ISNULL(h.estado,0) = 5 then 'TRABAJANDO' END as ESTADO,analista_3 as ANALISTA
from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 and c.ccia ='01'
LEFT JOIN HojaCotizacion H ON Q.ncom_3 = H.od+ H.od_version
where ncom_3 LIKE '03%' and fcom_3 >='20240928' and h.estado in (2,3)  and q.finped_3 ='0' ORDER BY  ncom_3 asc", conx)
        cmd.Fill(da1)
        DataGridView1.DataSource = da1
        DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(3).Width = 75
        DataGridView1.Columns(4).Width = 65
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = 213
        DataGridView1.Columns(7).Width = 115
        DataGridView1.Columns(8).Width = 213
        DataGridView1.Columns(11).Width = 62
        DataGridView1.Columns(12).Width = 73
        DataGridView1.Columns(14).Width = 80
        DataGridView1.Columns(15).Width = 95

    End Sub
    Private Sub mostrar_produccion()
        abrir()
        da1.Clear()
        DataGridView1.DataSource = ""
        Dim cmd As New SqlDataAdapter("select substring(ncom_3,1,9) as OD,substring(ncom_3,10,1) AS VERSION,program_3 AS PO, descri_3 AS PRODUCTO,telaprinc_3 AS TELA,mruta1_3 AS RUTA,estilo_3 as ESTILO,c.nomb_10 as CLIENTE,case when mone_3 = 0 then 'SOLES' ELSE 'DOLARES' END as MONEDA,isnull( h.id_cotizacion,'') as COTIZACION,tela1_3,case when ISNULL(h.estado,0) = 0 then 'PENDIENTE' when ISNULL(h.estado,0) = 1 then 'TRABAJANDO' when ISNULL(h.estado,0) = 2 then 'CERRADO' when ISNULL(h.estado,0) = 3 then 'TRABAJANDO' when ISNULL(h.estado,0) = 4 then 'PENDIENTE' when ISNULL(h.estado,0) = 5 then 'TRABAJANDO' END as ESTADO,analista_3 as ANALISTA
from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 and c.ccia ='01'
LEFT JOIN HojaCotizacion H ON Q.ncom_3 = H.od+ H.od_version
where ncom_3 LIKE '03%' and fcom_3 >='20240928' and h.estado in(4,5)  and q.finped_3 ='0' ORDER BY  ncom_3 asc", conx)
        cmd.Fill(da1)
        DataGridView1.DataSource = da1
        DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(3).Width = 75
        DataGridView1.Columns(4).Width = 65
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = 213
        DataGridView1.Columns(7).Width = 115
        DataGridView1.Columns(8).Width = 213
        DataGridView1.Columns(11).Width = 62
        DataGridView1.Columns(12).Width = 73
        DataGridView1.Columns(14).Width = 80
        DataGridView1.Columns(15).Width = 95
    End Sub
    Private Sub AlmacenCotizacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True
        ProgressBar1.Visible = False
        Label5.Visible = False
        validar()
    End Sub
    Sub validar()
        If Label3.Text.ToString().Trim() = "UDP" Then
            mostar_todos()
        Else
            If Label3.Text.ToString().Trim() = "PRODUCCION" Then
                mostrar_produccion()
            Else
                If Label3.Text.ToString().Trim() = "ALMACEN" Then
                    mostrar_compras()
                Else
                    If Label3.Text.ToString().Trim() = "COMERCIAL MANUFACTURA" Then
                        mostar_todos()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "UDP" Then
            If Label3.Text.ToString().Trim() = "UDP" Then
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
                If Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value) = "0" Then
                    cot.TextBox11.Text = "SOLES"
                Else
                    cot.TextBox11.Text = "DOLARES"
                End If
                If Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value).Length = 0 Then
                    Dim func As New FCOTIZACION
                    cot.TextBox10.Text = func.CORRELATIVO_COTIZACION() + 1
                    Select Case cot.TextBox10.Text.Length

                        Case "1" : cot.TextBox10.Text = "0000000" & "" & cot.TextBox10.Text
                        Case "2" : cot.TextBox10.Text = "000000" & "" & cot.TextBox10.Text
                        Case "3" : cot.TextBox10.Text = "00000" & "" & cot.TextBox10.Text
                        Case "4" : cot.TextBox10.Text = "0000" & "" & cot.TextBox10.Text
                        Case "5" : cot.TextBox10.Text = "000" & "" & cot.TextBox10.Text
                        Case "6" : cot.TextBox10.Text = "00" & "" & cot.TextBox10.Text
                        Case "7" : cot.TextBox10.Text = "0" & "" & cot.TextBox10.Text
                        Case "8" : cot.TextBox10.Text = cot.TextBox10.Text
                    End Select
                    cot.Label11.Text = "0"
                Else
                    cot.TextBox10.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value)
                    cot.Label11.Text = "1"
                End If
                cot.GroupBox7.Enabled = False
                cot.Label12.Text = "1"
                cot.TextBox10.Focus()
                SendKeys.Send("{ENTER}")
                cot.ShowDialog()
            Else
                MsgBox("Su Usuario no Tiene Permiso para ingresar a este Formulario")
            End If


        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "LOG" Then
            If (DataGridView1.Rows(e.RowIndex).Cells(12).Value.ToString().Trim().Length() > 0) Then
                If Label3.Text.ToString().Trim() = "ALMACEN" Then
                    Dim cot1 As New CotizacionUdp
                    cot1.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
                    cot1.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value)
                    cot1.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
                    cot1.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(9).Value)
                    cot1.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(8).Value)
                    cot1.TextBox6.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value)
                    cot1.TextBox8.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value)
                    cot1.TextBox9.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
                    cot1.TextBox14.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value)
                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value) = "0" Then
                        cot1.TextBox11.Text = "SOLES"
                    Else
                        cot1.TextBox11.Text = "DOLARES"
                    End If
                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value).Length = 0 Then
                        Dim func As New FCOTIZACION
                        cot1.TextBox10.Text = func.CORRELATIVO_COTIZACION() + 1
                        Select Case cot1.TextBox10.Text.Length

                            Case "1" : cot1.TextBox10.Text = "0000000" & "" & cot1.TextBox10.Text
                            Case "2" : cot1.TextBox10.Text = "000000" & "" & cot1.TextBox10.Text
                            Case "3" : cot1.TextBox10.Text = "00000" & "" & cot1.TextBox10.Text
                            Case "4" : cot1.TextBox10.Text = "0000" & "" & cot1.TextBox10.Text
                            Case "5" : cot1.TextBox10.Text = "000" & "" & cot1.TextBox10.Text
                            Case "6" : cot1.TextBox10.Text = "00" & "" & cot1.TextBox10.Text
                            Case "7" : cot1.TextBox10.Text = "0" & "" & cot1.TextBox10.Text
                            Case "8" : cot1.TextBox10.Text = cot1.TextBox10.Text
                        End Select
                    Else
                        cot1.TextBox10.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value)
                    End If
                    cot1.GroupBox7.Enabled = False
                    cot1.Label12.Text = "3"
                    cot1.TextBox10.Focus()
                    SendKeys.Send("{ENTER}")
                    cot1.ShowDialog()
                Else
                    MsgBox("Su Usuario no Tiene Permiso para ingresar a este Formulario")
                End If
            Else
                MsgBox("Esta Opcion no esta Habilitada hasta que Udp Ingrese Información")
            End If



        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PROD" Then
            If (DataGridView1.Rows(e.RowIndex).Cells(12).Value.ToString().Trim().Length() > 0) Then
                If Label3.Text.ToString().Trim() = "PRODUCCION" Then
                    Dim cot2 As New CotizacionUdp
                    cot2.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
                    cot2.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value)
                    cot2.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
                    cot2.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(9).Value)
                    cot2.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(8).Value)
                    cot2.TextBox6.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value)
                    cot2.TextBox8.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value)
                    cot2.TextBox9.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
                    cot2.TextBox14.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value)
                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value) = "0" Then
                        cot2.TextBox11.Text = "SOLES"
                    Else
                        cot2.TextBox11.Text = "DOLARES"
                    End If
                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value).Length = 0 Then
                        Dim func As New FCOTIZACION
                        cot2.TextBox10.Text = func.CORRELATIVO_COTIZACION() + 1
                        Select Case cot2.TextBox10.Text.Length

                            Case "1" : cot2.TextBox10.Text = "0000000" & "" & cot2.TextBox10.Text
                            Case "2" : cot2.TextBox10.Text = "000000" & "" & cot2.TextBox10.Text
                            Case "3" : cot2.TextBox10.Text = "00000" & "" & cot2.TextBox10.Text
                            Case "4" : cot2.TextBox10.Text = "0000" & "" & cot2.TextBox10.Text
                            Case "5" : cot2.TextBox10.Text = "000" & "" & cot2.TextBox10.Text
                            Case "6" : cot2.TextBox10.Text = "00" & "" & cot2.TextBox10.Text
                            Case "7" : cot2.TextBox10.Text = "0" & "" & cot2.TextBox10.Text
                            Case "8" : cot2.TextBox10.Text = cot2.TextBox10.Text
                        End Select
                    Else
                        cot2.TextBox10.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value)
                    End If
                    cot2.GroupBox2.Enabled = False
                    cot2.GroupBox3.Enabled = False
                    cot2.GroupBox4.Enabled = True
                    cot2.GroupBox5.Enabled = True
                    cot2.GroupBox6.Enabled = False
                    cot2.GroupBox7.Enabled = True
                    cot2.Label12.Text = "5"
                    cot2.TextBox10.Focus()
                    SendKeys.Send("{ENTER}")
                    cot2.ShowDialog()
                Else
                    MsgBox("Su Usuario no Tiene Permiso para ingresar a este Formulario")
                End If
            Else
                MsgBox("Esta Opcion no esta Habilitada hasta que Udp Ingrese Información")
            End If



        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OD" Then
            Od_Udp.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
            Od_Udp.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value)
            Od_Udp.ShowDialog()
        End If

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PO" Then
            ReportePm.TextBox1.Text = "01"
            ReportePm.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString())
            ReportePm.TextBox3.Text = "1"
            ReportePm.ShowDialog()
        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Columns(e.ColumnIndex).Name = "ESTADO" Then
            ' Verificar si el valor de la celda es "2"
            If Label3.Text.ToString().Trim() = "UDP" Then
                If DataGridView1.Rows(e.RowIndex).Cells(14).Value.ToString().Trim() = "PENDIENTE" Then
                    DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightGreen
                Else
                    If DataGridView1.Rows(e.RowIndex).Cells(14).Value.ToString().Trim() = "CERRADO" Then
                        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Black
                        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White

                    End If

                End If

            End If

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim export As New Exportar
        export.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click
        validar()
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da1.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "OD" & " like '%" & TextBox1.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(3).Width = 75
                DataGridView1.Columns(4).Width = 65
                DataGridView1.Columns(5).Width = 80
                DataGridView1.Columns(6).Width = 213
                DataGridView1.Columns(7).Width = 115
                DataGridView1.Columns(8).Width = 213
                DataGridView1.Columns(11).Width = 62
                DataGridView1.Columns(12).Width = 73
                DataGridView1.Columns(14).Width = 80
                DataGridView1.Columns(15).Width = 95
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
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da1.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "CLIENTE" & " like '%" & TextBox2.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(3).Width = 75
                DataGridView1.Columns(4).Width = 65
                DataGridView1.Columns(5).Width = 80
                DataGridView1.Columns(6).Width = 213
                DataGridView1.Columns(7).Width = 115
                DataGridView1.Columns(8).Width = 213
                DataGridView1.Columns(11).Width = 62
                DataGridView1.Columns(12).Width = 73
                DataGridView1.Columns(14).Width = 80
                DataGridView1.Columns(15).Width = 95
            Else
                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar2()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        mostar_todos_cerrado()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        mostar_todos()
    End Sub
End Class