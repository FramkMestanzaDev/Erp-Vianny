Imports System.Data.SqlClient

Public Class DetTelaProduccion
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Public comando As SqlCommand
    Dim da As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")


            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub DetTelaProduccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MOSTRAR()
    End Sub
    Sub MOSTRAR()
        abrir()
        da.Clear()
        DataGridView1.DataSource = Nothing
        Dim cmd As New SqlDataAdapter("SELECT ncom_3 AS OP, tela AS TELA, cg.nomb_17 AS PRODUCTO, dubica AS DETALLE, 
                                    ROUND(CASE WHEN cg.familia_17 = 'TPL' THEN cp.cxplineal * cantp_3 ELSE cp.kilos * q.cantp_3 END, 2) AS 'K/M', 
                                    CASE WHEN cg.familia_17 = 'TPL' THEN 'TELA PLANA' ELSE 'TELA PUNTO' END AS FAMILIA, 
                                    isnull((select sum(CanDtm) from DetalleTelaManufactura WHERE  OpDtm =cp.op and CopDtm =cp.tela),0) as DESPACHADO 
                                    FROM custom_vianny.dbo.Consumo_Op_DET cp
                                    LEFT JOIN custom_vianny.dbo.cag1700 cg ON cp.ccia = cg.ccia AND cp.tela = cg.linea_17 
                                    LEFT JOIN custom_vianny.dbo.qag0300 q ON cp.ccia = q.ccia AND cp.op = q.ncom_3 
                                    WHERE cp.ccia = '" & Label3.Text & "' AND op = '" & Label2.Text & "'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(2).Width = 320
        DataGridView1.Columns(0).Width = 140
        DataGridView1.Columns(1).Width = 140
        'DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(0).Visible = False
        'For i As Integer = 0 To DataGridView1.Rows.Count - 1

        '    If Trim(DataGridView1.Rows(i).Cells(6).Value.ToString()) = "SELECCIONADO" Then
        '        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
        '    End If
        'Next

        conx.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            Dim lotes As New ProdLotes
            ' Accede a las celdas de la fila seleccionada, por ejemplo:
            Dim valorCeldadespachado As Double = Convert.ToDouble(selectedRow.Cells(6).Value)
            Dim valorCeldasolicitado As Double = Convert.ToDouble(selectedRow.Cells(4).Value)
            Dim rowIndex As Integer = DataGridView1.SelectedRows(0).Index
            If valorCeldadespachado < valorCeldasolicitado Then
                lotes.Label1.Text = Label3.Text
                lotes.Label2.Text = Modulo_Almacen.Label8.Text
                lotes.Label4.Text = Trim(selectedRow.Cells(1).Value.ToString)
                lotes.Label5.Text = Trim(selectedRow.Cells(5).Value.ToString)
                lotes.Label7.Text = Label2.Text
                lotes.Label11.Text = Label5.Text
                lotes.Label10.Text = selectedRow.Cells(3).Value.ToString.Trim
                lotes.TextBox1.Text = Convert.ToDouble(selectedRow.Cells(4).Value.ToString) - Convert.ToDouble(selectedRow.Cells(6).Value.ToString)
                lotes.ShowDialog()

                MOSTRAR()
            Else
                MsgBox("NO SE PUEDE AGREGAR DE NUEVO EL ITEMS YA ESTA SELECCIONADO, SI NECESITAR QUITAR LA SELECCION HAGA CLIC EN EL BOTON DESELECCIONAR")
            End If
        End If

    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RptTelaOp.TextBox1.Text = Label3.Text
        RptTelaOp.TextBox2.Text = Label2.Text
        RptTelaOp.ShowDialog()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        'If DataGridView1.SelectedRows.Count > 0 Then
        '    Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

        '    ' Accede a las celdas de la fila seleccionada, por ejemplo:
        '    Dim valorCeldaPorIndice As String = selectedRow.Cells(6).Value.ToString
        '    Dim rowIndex As Integer = DataGridView1.SelectedRows(0).Index
        '    If Trim(valorCeldaPorIndice) = "SELECCIONADO" Then
        '        abrir()
        '        Dim cmd20 As New SqlCommand("delete from DetalleTelaManufactura  where OpDtm =@op and CopDtm =@linea and CodDtm =@codigo", conx)
        '        cmd20.Parameters.AddWithValue("@op", Trim(selectedRow.Cells(0).Value.ToString))
        '        cmd20.Parameters.AddWithValue("@linea", Trim(selectedRow.Cells(1).Value.ToString))
        '        cmd20.Parameters.AddWithValue("@codigo", Trim(selectedRow.Cells(7).Value.ToString))
        '        cmd20.ExecuteNonQuery()
        '        MsgBox("Se Actualizo la Informacion Correctamnete")
        '        MOSTRAR()
        '    Else
        '        If Trim(valorCeldaPorIndice) = "DESPACHADO" Then
        '            MsgBox("EL ITEMS ESTA EN ESTADO DESPACHADO NO SE PUEDE EDITAR")
        '        Else
        '            MsgBox("EL ITEMS ESTA   EN ESTADO PENDIENTE NO ES NECESARIA ESTA OPCION")
        '        End If

        '    End If
        'End If
    End Sub

    Private Sub DetTelaProduccion_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim suma As Integer
        For i = 0 To DataGridView1.Rows.Count - 1
            If Convert.ToDouble(DataGridView1.Rows(i).Cells(6).Value) > 0 Then
                suma = suma + 1
            End If
        Next
        If suma > 0 Then
            ReqTelaProduccion.DataGridView1.Rows(Label4.Text).Cells(1).Value = True
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        EnviarCorreo.TextBox1.Text = ""
        EnviarCorreo.TextBox2.Text = ""
        EnviarCorreo.TextBox1.Text = "Para Informarles que el consumo de tela de la op" + Label2.Text.ToString().Trim() + " no esta Ingresado, Realizar el ingreso para despachar al Area de Corte"
        EnviarCorreo.Label3.Text = Modulo_Almacen.Label6.Text.ToString().Trim()
        EnviarCorreo.Label5.Text = "0"
        EnviarCorreo.Label4.Text = Label2.Text.ToString().Trim()
        EnviarCorreo.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        EnviarCorreo.TextBox1.Text = ""
        EnviarCorreo.TextBox2.Text = ""
        Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
        EnviarCorreo.TextBox1.Text = "Para Informarles que no se tiene Stock del producto " + filaSeleccionada.Cells("TELA").Value.ToString().Trim() + " Producto : " + filaSeleccionada.Cells("PRODUCTO").Value.ToString().Trim() + " , Validar que sea el Producto correcto para despachar a Corte"
        EnviarCorreo.Label3.Text = Modulo_Almacen.Label6.Text.ToString().Trim()
        EnviarCorreo.Label4.Text = Label2.Text.ToString().Trim()
        EnviarCorreo.Label5.Text = "0"
        EnviarCorreo.ShowDialog()
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(6).Value) >= Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(4).Value) Then
            ' Verificar si el valor de la celda es "2"
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Black
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White
            DataGridView1.Rows(e.RowIndex).ReadOnly = True
        End If
    End Sub
End Class