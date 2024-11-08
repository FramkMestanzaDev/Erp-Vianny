Imports System.Data.SqlClient
Public Class StatusProduccionPP
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim da As New DataTable
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
        abrir()
        da.Clear()
        Dim cmd As New SqlDataAdapter("exec Sp_StatusProduccionGraficos1 '" + Trim(Label2.Text) + "'", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 85
        DataGridView1.Columns(1).Width = 89
        DataGridView1.Columns(2).Width = 200
        DataGridView1.Columns(3).Width = 65
        DataGridView1.Columns(4).Width = 65
        DataGridView1.Columns(5).Width = 200
        DataGridView1.Columns(6).Width = 90
        DataGridView1.Columns(7).Width = 87
        DataGridView1.Columns(8).Width = 87
        DataGridView1.Columns(9).Width = 87
        DataGridView1.Columns(10).Width = 87
        DataGridView1.Columns(11).Width = 87
        DataGridView1.Columns(12).Width = 87
        DataGridView1.Columns(13).Width = 87

        DataGridView1.Columns(0).Frozen = True
        DataGridView1.Columns(1).Frozen = True
        DataGridView1.Columns(2).Frozen = True
        DataGridView1.Columns(3).Frozen = True

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "Op" & " like '%" & TextBox2.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 85
                DataGridView1.Columns(1).Width = 89
                DataGridView1.Columns(2).Width = 200
                DataGridView1.Columns(3).Width = 65
                DataGridView1.Columns(4).Width = 65
                DataGridView1.Columns(5).Width = 200
                DataGridView1.Columns(6).Width = 90
                DataGridView1.Columns(7).Width = 87
                DataGridView1.Columns(8).Width = 87
                DataGridView1.Columns(9).Width = 87
                DataGridView1.Columns(10).Width = 87
                DataGridView1.Columns(11).Width = 87
                DataGridView1.Columns(12).Width = 87
                DataGridView1.Columns(13).Width = 87

                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True

            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub buscarPo()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "Po" & " like '%" & TextBox3.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 85
                DataGridView1.Columns(1).Width = 89
                DataGridView1.Columns(2).Width = 200
                DataGridView1.Columns(3).Width = 65
                DataGridView1.Columns(4).Width = 65
                DataGridView1.Columns(5).Width = 200
                DataGridView1.Columns(6).Width = 90
                DataGridView1.Columns(7).Width = 87
                DataGridView1.Columns(8).Width = 87
                DataGridView1.Columns(9).Width = 87
                DataGridView1.Columns(10).Width = 87
                DataGridView1.Columns(11).Width = 87
                DataGridView1.Columns(12).Width = 87
                DataGridView1.Columns(13).Width = 87

                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True

            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub StatusProduccionPP_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(5, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(5, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(6, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(6, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(7, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(8).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(8, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(8, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(9).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(9, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(9, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(10, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(11, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(11, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(12).Value.ToString()) = "PENDIENTE" Then
            DataGridView1.Item(12, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(12, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(13).Value.ToString()).Length = 0 Then
            DataGridView1.Rows(e.RowIndex).Cells(13).Value = "PENDIENTE"
            DataGridView1.Item(13, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(13, e.RowIndex).Style.BackColor = Color.LightGreen
            DataGridView1.Item(0, e.RowIndex).Style.ForeColor = Color.Black
            DataGridView1.Item(0, e.RowIndex).Style.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Produccion" Then
            Status_op_simple.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
            Status_op_simple.TextBox2.Text = Trim(Label2.Text)
            Status_op_simple.ShowDialog()
        End If
        If e.ColumnIndex = 0 Or e.ColumnIndex = 1 Then
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL REPORTE LISTADO OP x PO?  -- PO:  " + DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim(), "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                ListadoOpxPo.TextBox1.Text = Label2.Text.ToString().Trim()
                ListadoOpxPo.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
                ListadoOpxPo.ShowDialog()
            End If

        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Matriz" Then
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL REPORTE MATRIZ DE LA OP? " + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim(), "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                ReporteMatriz.TextBox1.Text = Label2.Text.ToString().Trim()
                ReporteMatriz.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                ReporteMatriz.ShowDialog()
            End If

        End If

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "ConsumoTela" Then

            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL REPORTE CONSUMO DE TELA DE LA PO?  PO --- " + DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim(), "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Reporte_Consmmo.TextBox1.Text = Label2.Text.ToString().Trim()
                Reporte_Consmmo.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
                Reporte_Consmmo.ShowDialog()
            End If

        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OcTela" Then
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL REPORTE ORDEN DE COMPRA DE TELA DE LA PO?   " + DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim(), "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Orden_Compra_Ac.Label3.Text = Label2.Text.ToString().Trim()
                Orden_Compra_Ac.CheckBox1.Checked = True
                Orden_Compra_Ac.CheckBox4.Checked = True
                Orden_Compra_Ac.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
                Orden_Compra_Ac.TextBox4.Text = "08"
                Orden_Compra_Ac.TextBox5.Text = "TELA"
                Orden_Compra_Ac.Show(Me)
            End If
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OcAv" Then
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL REPORTE ORDEN DE COMPRA DE AVIOS DE LA PO?   " + DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim(), "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Orden_Compra_Ac.Label3.Text = Label2.Text.ToString().Trim()
                Orden_Compra_Ac.CheckBox1.Checked = True
                Orden_Compra_Ac.CheckBox4.Checked = True
                Orden_Compra_Ac.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
                Orden_Compra_Ac.TextBox4.Text = "13"
                Orden_Compra_Ac.TextBox5.Text = "AVIOS"
                Orden_Compra_Ac.Show(Me)
            End If
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "IngresoTela" Then
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL REPORTE PARA VER LOS INGRESOS DE TELA DE LA PO?   " + DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim(), "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Reporte_Ingreso_Salida.TextBox4.Text = Label5.Text
                Reporte_Ingreso_Salida.TextBox5.Text = 1
                Reporte_Ingreso_Salida.Text = "REPORTE PARTE INGRESO"
                Reporte_Ingreso_Salida.Label4.Text = Label2.Text.ToString().Trim()
                Reporte_Ingreso_Salida.TextBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
                Reporte_Ingreso_Salida.GroupBox1.Visible = False
                Reporte_Ingreso_Salida.GroupBox2.Visible = False
                Reporte_Ingreso_Salida.GroupBox3.Visible = False
                Reporte_Ingreso_Salida.Button4.Enabled = True
                Reporte_Ingreso_Salida.Show(Me)
            End If
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Molde" Then
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("ESTAS INGRESANDO AL VISUALIZAR LA INFORMACION DEL MOLDE?   " + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim(), "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                ObsMolde.Label1.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                ObsMolde.Show(Me)
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ext As New ExportaruDP
        ext.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
        Dim po As String = filaSeleccionada.Cells("Po").Value.ToString().Trim()
        ConsultarConsumo.Label3.Text = Label2.Text.ToString().Trim()
        If Label2.Text.ToString().Trim() = "01" Then
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

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscarPo()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class