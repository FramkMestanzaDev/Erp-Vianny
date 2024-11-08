Imports System.Data.SqlClient
Public Class HojaContizacionAlm
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
        Dim cmd As New SqlDataAdapter("select substring(ncom_3,1,9) as OD,substring(ncom_3,10,1) AS VERSION,program_3 AS PO, descri_3 AS PRODUCTO,telaprinc_3 AS TELA,mruta1_3 AS RUTA,estilo_3 as ESTILO,c.nomb_10 as CLIENTE,case when mone_3 = 0 then 'SOLES' ELSE 'DOLARES' END as MONEDA,isnull( h.id_cotizacion,'') as COTIZACION,tela1_3,ISNULL(h.estado,0) as ESTADO
from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 and c.ccia ='01'
LEFT JOIN HojaCotizacion H ON Q.ncom_3 = H.od+ H.od_version
where ncom_3 LIKE '03%' and fcom_3 >='20240928' ORDER BY  ncom_3 asc", conx)
        cmd.Fill(da1)
        DataGridView1.DataSource = da1
        DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(3).Width = 78
        DataGridView1.Columns(4).Width = 70
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = 215
        DataGridView1.Columns(7).Width = 117
        DataGridView1.Columns(8).Width = 215

    End Sub
    Private Sub mostrar_compras()
        abrir()
        da1.Clear()
        DataGridView1.DataSource = ""
        Dim cmd As New SqlDataAdapter("select substring(ncom_3,1,9) as OD,substring(ncom_3,10,1) AS VERSION,program_3 AS PO, descri_3 AS PRODUCTO,telaprinc_3 AS TELA,mruta1_3 AS RUTA,estilo_3 as ESTILO,c.nomb_10 as CLIENTE,case when mone_3 = 0 then 'SOLES' ELSE 'DOLARES' END as MONEDA,isnull( h.id_cotizacion,'') as COTIZACION,tela1_3,ISNULL(h.estado,0) as ESTADO
from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 and c.ccia ='01'
LEFT JOIN HojaCotizacion H ON Q.ncom_3 = H.od+ H.od_version
where ncom_3 LIKE '03%' and fcom_3 >='20240928' and h.estado in (2,3) ORDER BY  ncom_3 asc", conx)
        cmd.Fill(da1)
        DataGridView1.DataSource = da1
        DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(3).Width = 78
        DataGridView1.Columns(4).Width = 70
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = 215
        DataGridView1.Columns(7).Width = 117
        DataGridView1.Columns(8).Width = 215

    End Sub
    Private Sub mostrar_produccion()
        abrir()
        da1.Clear()
        DataGridView1.DataSource = ""
        Dim cmd As New SqlDataAdapter("select substring(ncom_3,1,9) as OD,substring(ncom_3,10,1) AS VERSION,program_3 AS PO, descri_3 AS PRODUCTO,telaprinc_3 AS TELA,mruta1_3 AS RUTA,estilo_3 as ESTILO,c.nomb_10 as CLIENTE,case when mone_3 = 0 then 'SOLES' ELSE 'DOLARES' END as MONEDA,isnull( h.id_cotizacion,'') as COTIZACION,tela1_3,ISNULL(h.estado,0)as ESTADO
from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 and c.ccia ='01'
LEFT JOIN HojaCotizacion H ON Q.ncom_3 = H.od+ H.od_version
where ncom_3 LIKE '03%' and fcom_3 >='20240928' and h.estado in(4,5) ORDER BY  ncom_3 asc", conx)
        cmd.Fill(da1)
        DataGridView1.DataSource = da1
        DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(3).Width = 78
        DataGridView1.Columns(4).Width = 70
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = 215
        DataGridView1.Columns(7).Width = 117
        DataGridView1.Columns(8).Width = 215

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
    Private Sub HojaContizacionAlm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        validar()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        CotizacionUdp.limpiarData()

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
                    cot2.ShowDialog()
                Else
                    MsgBox("Su Usuario no Tiene Permiso para ingresar a este Formulario")
                End If
            Else
                MsgBox("Esta Opcion no esta Habilitada hasta que Udp Ingrese Información")
            End If



        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        ' Verificar que la columna "VERSION" (índice 2) sea la columna correcta
        If DataGridView1.Columns(e.ColumnIndex).Name = "ESTADO" Then
            ' Verificar si el valor de la celda es "2"
            If Label3.Text.ToString().Trim() = "UDP" Then
                If e.Value IsNot Nothing AndAlso e.Value.ToString() = "0" Then
                    DataGridView1.Rows(e.RowIndex).Cells(14).Value = "PENDIENTE"
                    DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightGreen
                Else
                    If e.Value IsNot Nothing AndAlso e.Value.ToString() = "2" Then
                        DataGridView1.Rows(e.RowIndex).Cells(14).Value = "CERRADO"
                        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(14).Value = "TRABAJANDO"
                    End If

                End If
            Else
                If Label3.Text.ToString().Trim() = "PRODUCCION" Then
                    If e.Value IsNot Nothing AndAlso e.Value.ToString() = "4" Then
                        DataGridView1.Rows(e.RowIndex).Cells(14).Value = "PENDIENTE"
                        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightGreen
                    Else
                        If e.Value IsNot Nothing AndAlso e.Value.ToString() = "5" Then
                            DataGridView1.Rows(e.RowIndex).Cells(14).Value = "TRABAJANDO"

                        Else
                            DataGridView1.Rows(e.RowIndex).Cells(14).Value = "CERRADO"
                            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Black
                            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White
                        End If
                    End If
                Else
                    If Label3.Text.ToString().Trim() = "ALMACEN" Then
                        If e.Value IsNot Nothing AndAlso e.Value.ToString() = "2" Then
                            DataGridView1.Rows(e.RowIndex).Cells(14).Value = "PENDIENTE"
                            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightGreen
                        Else
                            If e.Value IsNot Nothing AndAlso e.Value.ToString() = "3" Then
                                DataGridView1.Rows(e.RowIndex).Cells(14).Value = "TRABAJANDO"
                            Else
                                DataGridView1.Rows(e.RowIndex).Cells(14).Value = "CERRADO"
                                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Black
                                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White
                            End If
                        End If

                    End If
                End If
            End If

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim export As New Exportar
        export.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        validar()
    End Sub
End Class