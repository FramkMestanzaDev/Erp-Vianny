Imports System.Data.SqlClient
Public Class Pre_produccion
    Public conx As SqlConnection
    Public comando As SqlCommand
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select  q.program_3 as PO,	ncom_3 as OP,c.nomb_10 as CLIENTE, descri_3 AS PRODUCTO,estilo_3 AS ESTILO,um.usuario as VENDEDOR,Q.cantp_3 AS PROGRAMADO, (select case when COUNT(tela) = 0 then 'PENDIENTE' else 'INGRESADO'  END
 from custom_vianny.dbo.Consumo_Op_DET where ccia = q.ccia and op =q.ncom_3) as 'CONSUMO TELA', isnull(CAST( rt.FEnRtp as varchar(12)),'PENDIENTE') AS 'REQ. PROD' ,case when LTRIM(RTRIM(CAST( mruta7_3 AS VARCHAR(MAX)))) = '' then 'PENDIENTE' ELSE rtrim(ltrim(CAST( mruta7_3 AS VARCHAR(MAX)))) END 
  AS MOLDE,ISNULL( C1.nomb_10,'PENDIENTE') as MODELISTA,
case when (select case when count(ncom_8) >0 then '2' end from custom_vianny.dbo.qamc800 where ncom_8 =q.ncom_3 and ccia_8 =q.ccia) ='2' then 'GENERADO'
ELSE 'PENDIENTE'
end as EXPLOSION_UDP,
case when cierrematcon_3 = 1 then 'VALIDADO LOGIS.' ELSE 'PENDIENTE'END  AS EXPLOCION_AVIOS
from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3=c.fich_10 
LEFT join custom_vianny.dbo.cag1000 c1 on  q.ccia = c1.ccia and q.modelista_3=c1.ruc_10 and c1.tdoc_10='01' AND Q.modelista_3 <> ''
LEFT JOIN RequerimientoTelaProd rt on rt.OpRtp = q.ncom_3 and rt.CanRtp = 0 
left join usuario_modulo um on q.merchan_3 = um.merchan_3
WHERE  YEAR(Q.fcom_3) = '" + Trim(TextBox1.Text) + "'  and q.ccia ='01' and broker_3 not in ( 0003,0000) and ncom_3 like '01%' and flag_3 ='1' and finped_3 ='0'  ORDER BY q.ncom_3 desc", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 87
        DataGridView1.Columns(1).Width = 87
        DataGridView1.Columns(2).Width = 200
        DataGridView1.Columns(3).Width = 250
        DataGridView1.Columns(4).Width = 100
        DataGridView1.Columns(5).Width = 90
        DataGridView1.Columns(6).Width = 87
        DataGridView1.Columns(7).Width = 87
        DataGridView1.Columns(8).Width = 90
        DataGridView1.Columns(9).Width = 220
        DataGridView1.Columns(10).Width = 180
        DataGridView1.Columns(11).Width = 100
        DataGridView1.Columns(12).Width = 95
        DataGridView1.Columns(0).Frozen = True
        DataGridView1.Columns(1).Frozen = True
        DataGridView1.Columns(2).Frozen = True
        DataGridView1.Columns(3).Frozen = True
        Dim o As Integer
        o = DataGridView1.Rows.Count


    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "OP" & " like '%" & TextBox2.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 87
                DataGridView1.Columns(1).Width = 87
                DataGridView1.Columns(2).Width = 200
                DataGridView1.Columns(3).Width = 250
                DataGridView1.Columns(4).Width = 100
                DataGridView1.Columns(5).Width = 90
                DataGridView1.Columns(6).Width = 87
                DataGridView1.Columns(7).Width = 87
                DataGridView1.Columns(8).Width = 90
                DataGridView1.Columns(9).Width = 220
                DataGridView1.Columns(10).Width = 180
                DataGridView1.Columns(11).Width = 100
                DataGridView1.Columns(12).Width = 95
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
            Else
                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(0).Width = 87
                DataGridView1.Columns(1).Width = 87
                DataGridView1.Columns(2).Width = 200
                DataGridView1.Columns(3).Width = 250
                DataGridView1.Columns(4).Width = 100
                DataGridView1.Columns(5).Width = 90
                DataGridView1.Columns(6).Width = 87
                DataGridView1.Columns(7).Width = 87
                DataGridView1.Columns(8).Width = 90
                DataGridView1.Columns(9).Width = 220
                DataGridView1.Columns(10).Width = 180
                DataGridView1.Columns(11).Width = 100
                DataGridView1.Columns(12).Width = 95
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                Dim o As Integer
                o = DataGridView1.Rows.Count


            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "ESTILO" & " like '%" & TextBox4.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(0).Width = 87
                DataGridView1.Columns(1).Width = 87
                DataGridView1.Columns(2).Width = 200
                DataGridView1.Columns(3).Width = 250
                DataGridView1.Columns(4).Width = 100
                DataGridView1.Columns(5).Width = 90
                DataGridView1.Columns(6).Width = 87
                DataGridView1.Columns(7).Width = 87
                DataGridView1.Columns(8).Width = 90
                DataGridView1.Columns(9).Width = 220
                DataGridView1.Columns(10).Width = 180
                DataGridView1.Columns(11).Width = 100
                DataGridView1.Columns(12).Width = 95
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                Dim o As Integer
                o = DataGridView1.Rows.Count


            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
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

    Private Sub Pre_produccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(6).Value) = 0 Then
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(1, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(1, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(2, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(2, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(3, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(3, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(6, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(6, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If DataGridView1.Rows(e.RowIndex).Cells(7).Value.ToString().Trim() = "PENDIENTE" Then
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(1, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(1, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(2, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(2, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(3, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(3, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(7, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.LightGreen
        End If

        If Trim(DataGridView1.Rows(e.RowIndex).Cells(8).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(1, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(1, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(2, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(2, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(3, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(3, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(8, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(8, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If DataGridView1.Rows(e.RowIndex).Cells(9).Value.ToString().Trim() = "PENDIENTE" Then
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(1, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(1, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(2, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(2, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(3, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(3, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(9, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(9, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(1, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(1, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(2, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(2, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(3, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(3, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(10, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(1, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(1, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(2, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(2, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(3, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(3, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(11, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(11, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(1, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(1, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(2, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(2, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(3, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(3, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(12, e.RowIndex).Style.ForeColor = Color.DarkRed
            DataGridView1.Item(12, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.ColumnIndex = 0 Or e.ColumnIndex = 1 Then
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL REPORTE LISTADO OP x PO?  -- PO:  " + DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim(), "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                ListadoOpxPo.TextBox1.Text = "01"
                ListadoOpxPo.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
                ListadoOpxPo.ShowDialog()
            End If

        End If
            If e.ColumnIndex = 9 Or e.ColumnIndex = 10 Then

            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL MODULO DE ASIGANCION DE MOLDE - OP?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Asignacion_Molde.Label9.Text = MDIParent2.Label10.Text.ToString().Trim()
                Asignacion_Molde.Label13.Text = MDIParent2.Label6.Text.ToString().Trim()
                Asignacion_Molde.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                Asignacion_Molde.Label18.Text = "1"
                Asignacion_Molde.TextBox2.Focus()
                SendKeys.Send("{ENTER}")
                Asignacion_Molde.Show()
            End If

        End If
        If e.ColumnIndex = 11 Then

            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL MODULO MATRIZ DE AVIOS - OP?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

            End If
            Matriz_Avios.Label10.Text = MDIParent2.Label10.Text.ToString().Trim()
            Matriz_Avios.Label11.Text = MDIParent2.Label8.Text.ToString().Trim()
            Matriz_Avios.TextBox9.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
            Matriz_Avios.Label16.Text = "1"
            Matriz_Avios.TextBox9.Focus()
            SendKeys.Send("{ENTER}")
            Matriz_Avios.Show(Me)
        End If
        If e.ColumnIndex = 7 Then

            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL REPORTE CONSUMO DE TELA DE LA PO?  PO --- " + DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim(), "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Reporte_Consmmo.TextBox1.Text = "01"
                Reporte_Consmmo.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
                Reporte_Consmmo.Show()
            End If

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
        Dim po As String = filaSeleccionada.Cells("PO").Value.ToString().Trim()
        ConsultarConsumo.Label3.Text = Label5.Text
        If Label5.Text = "01" Then
            ConsultarConsumo.TextBox1.Text = "VIANNY"
        Else
            ConsultarConsumo.TextBox1.Text = "GRAUS"
        End If
        ConsultarConsumo.CheckBox2.Checked = True
        ConsultarConsumo.TextBox2.Text = po
        If Not ConsultarConsumo.Visible Then
            ConsultarConsumo.Show(Me)
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ext As New ExportaruDP
        ext.llenarExcel(DataGridView1)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class