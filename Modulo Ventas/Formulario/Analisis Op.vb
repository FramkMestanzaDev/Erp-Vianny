Imports System.Data.SqlClient
Public Class Analisis_Op
    Public conx As SqlConnection
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Trim(TextBox1.Text).Length
                Case "1" : TextBox1.Text = "01" & "0000000" & TextBox1.Text
                Case "2" : TextBox1.Text = "01" & "000000" & TextBox1.Text
                Case "3" : TextBox1.Text = "01" & "00000" & TextBox1.Text
                Case "4" : TextBox1.Text = "01" & "0000" & TextBox1.Text
                Case "5" : TextBox1.Text = "01" & "000" & TextBox1.Text
                Case "6" : TextBox1.Text = "01" & "00" & TextBox1.Text
                Case "7" : TextBox1.Text = "01" & "0" & TextBox1.Text
                Case "8" : TextBox1.Text = "01" & TextBox1.Text
                Case "9" : TextBox1.Text = "0" & TextBox1.Text
            End Select
            TextBox4.Select()
        End If

    End Sub
    Dim da As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = ""
        da.Clear()
        abrir()
        Select Case Trim(TextBox1.Text).Length
            Case "1" : TextBox1.Text = "01" & "0000000" & TextBox1.Text
            Case "2" : TextBox1.Text = "01" & "000000" & TextBox1.Text
            Case "3" : TextBox1.Text = "01" & "00000" & TextBox1.Text
            Case "4" : TextBox1.Text = "01" & "0000" & TextBox1.Text
            Case "5" : TextBox1.Text = "01" & "000" & TextBox1.Text
            Case "6" : TextBox1.Text = "01" & "00" & TextBox1.Text
            Case "7" : TextBox1.Text = "01" & "0" & TextBox1.Text
            Case "8" : TextBox1.Text = "01" & TextBox1.Text
            Case "9" : TextBox1.Text = "0" & TextBox1.Text
        End Select
        Select Case Trim(TextBox4.Text).Length
            Case "1" : TextBox4.Text = "01" & "0000000" & TextBox4.Text
            Case "2" : TextBox4.Text = "01" & "000000" & TextBox4.Text
            Case "3" : TextBox4.Text = "01" & "00000" & TextBox4.Text
            Case "4" : TextBox4.Text = "01" & "0000" & TextBox4.Text
            Case "5" : TextBox4.Text = "01" & "000" & TextBox4.Text
            Case "6" : TextBox4.Text = "01" & "00" & TextBox4.Text
            Case "7" : TextBox4.Text = "01" & "0" & TextBox4.Text
            Case "8" : TextBox4.Text = "01" & TextBox4.Text
            Case "9" : TextBox1.Text = "0" & TextBox1.Text
        End Select
        Dim cmd As New SqlDataAdapter("select case when  talm_3a ='0300' then 'CORTE' when  talm_3a ='0400' then 'CONFECCION' when  talm_3a ='0500' then 'ACABADOS' when  talm_3a ='0600' then 'LAVANDERIA' when  talm_3a ='0700' then 'APLICACIONES Y PIEZAS' END AS ALMACEN ,CASE WHEN ccom_3a = 1 THEN 'INGRESO' ELSE 'SALIDA' END AS ORIGEN,qg.fich_3 as RUC , c.nomb_10 AS TALLER,CASE WHEN ccom_3a = 1 THEN  e.nomb_21e else ee.nomb_21s end AS MOTIVO, cASE WHEN ccom_3a = 1 then ncom_3a else ncom_3a + '/'+ grm_3 end AS NIA_NSA,
           cantk_3a AS CANTIDAD,pedido_3a AS OP,ordens_3a AS ORDEN_SERVICIO,ISNULL( p.preun1_4a,0) as PRECIO_S, ISNULL(P.preun2_4a,0) AS PRECIO_D, CASE WHEN flag_3a = '1' then 'ACTIVO' ELSE 'ANULADO' END AS SITUACION from custom_vianny.dbo.Qap3000 qp 
           inner join custom_vianny.dbo.Qag3000 qg on qp.Empr_3a = qg.Empr_3 and qp.Ano_3a =qg.Ano_3 and qp.talm_3a =qg.talm_3 and qp.ccom_3a = qg.ccom_3 and qp.ncom_3a=qg.ncom_3 
           left join custom_vianny.dbo.cag1000 c on  qg.Empr_3 = c.ccia and qg.fich_3 = c.fich_10
             left join custom_vianny.dbo.Qag210e e on qg.Empr_3 =e.Empr_21e and qg.orig_3 = e.dest_21e and qg.talm_3 = e.almac_21e
            left join custom_vianny.dbo.Qag210s ee on qg.Empr_3 =ee.Empr_21s and qg.orig_3 = ee.dest_21s and qg.talm_3 = ee.almac_21s
        left join custom_vianny.dbo.lap0400 p on qp.ordens_3a = p.ncom_4a and qp.Empr_3a = p.ccia_4a and qp.pedido_3a = p.pedido_4a
           where (pedido_3a >='" + Trim(TextBox1.Text) + "' and pedido_3a <='" + Trim(TextBox4.Text) + "')  and flag_3a ='1'  and Empr_3a ='" + Trim(TextBox3.Text) + "'
           group by  talm_3a,qg.subfase_3,ccom_3a,ncom_3a,cantk_3a,pedido_3a,ordens_3a,qg.fich_3,c.nomb_10,e.nomb_21e,p.preun1_4a,P.preun2_4a,qg.grm_3,qg.orig_3,ee.nomb_21s,flag_3a", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
        Else
            DataGridView1.DataSource = Nothing
        End If

        Dim exportarr As New Exportar
        exportarr.llenarExcel(DataGridView1)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Trim(TextBox4.Text).Length
                Case "1" : TextBox4.Text = "01" & "0000000" & TextBox4.Text
                Case "2" : TextBox4.Text = "01" & "000000" & TextBox4.Text
                Case "3" : TextBox4.Text = "01" & "00000" & TextBox4.Text
                Case "4" : TextBox4.Text = "01" & "0000" & TextBox4.Text
                Case "5" : TextBox4.Text = "01" & "000" & TextBox4.Text
                Case "6" : TextBox4.Text = "01" & "00" & TextBox4.Text
                Case "7" : TextBox4.Text = "01" & "0" & TextBox4.Text
                Case "8" : TextBox4.Text = "01" & TextBox4.Text

            End Select
            Button1.Select()
        End If
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
End Class