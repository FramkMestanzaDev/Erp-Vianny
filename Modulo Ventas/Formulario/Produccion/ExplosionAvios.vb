Imports System.Data.SqlClient
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO
Public Class ExplosionAvios
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim da, da1, da2, da3 As New DataTable
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

        If TextBox1.Text.ToString.Trim() <> "" Then
            Dim respuesta As DialogResult

            respuesta = MessageBox.Show("Se va a Precesar Todas la O.P. de la PO " + TextBox1.Text.ToString().Trim() + " si existe informacion esta sera sobreescrita?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                mostrarConsolidado()
                mostrarDetallado()
                TextBox1.Enabled = False
                Button1.Enabled = False
                Button2.Enabled = False
                Button4.Enabled = True
                TextBox2.Text = ""
                TextBox2.Enabled = True
            End If

        Else
                MsgBox("El Campo Po esta Vacio")
        End If

    End Sub
    Sub mostrarConsolidado()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("exec ExplosionAviosPo '" + TextBox1.Text.ToString().Trim() + "','" + Label3.Text.ToString().Trim() + "'", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView2.DataSource = da
            DataGridView2.Columns(0).Width = 140
            DataGridView2.Columns(1).Width = 300
            DataGridView2.Columns(2).Width = 45
            DataGridView2.Columns(3).Width = 75
            DataGridView2.Columns(4).Width = 75
            DataGridView2.Columns(5).Width = 75
            DataGridView2.Columns(6).Width = 75
            DataGridView2.Columns(7).Width = 75
            DataGridView2.Columns(8).Width = 75
            DataGridView2.Columns(10).Width = 45
            DataGridView2.Columns(9).Visible = False
            DataGridView2.Columns(0).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(0).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(1).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(1).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(2).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(2).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(3).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(3).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(4).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(4).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(6).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(6).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(8).DefaultCellStyle.BackColor = Color.DarkSlateBlue
            DataGridView2.Columns(8).DefaultCellStyle.ForeColor = Color.White
        Else
            DataGridView2.DataSource = Nothing
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            ListaPO.Label3.Text = "6"
            ListaPO.Show(Me)
        Else
            If e.KeyCode = Keys.Enter Then
                abrir()
                Dim Rsr19 As SqlDataReader
                Dim num As Integer = 0
                Dim sql101 As String = "SELECT COUNT(po_mcr) FROM custom_vianny.dbo.Requerimiento_Logistico_Resumen  A  Where CCia_mcr='" + Label3.Text.ToString().Trim() + "' AND po_mcr='" + TextBox1.Text.ToString().Trim() + "'"
                Dim cmd101 As New SqlCommand(sql101, conx)
                Rsr19 = cmd101.ExecuteReader()
                If Rsr19.Read() Then
                    num = Rsr19(0)

                End If
                Rsr19.Close()
                If num > 0 Then
                    mostrarInformacionTotal()
                    mostrarInformacionDetalle()
                    bloquear()
                Else
                    Button1.Enabled = True

                End If
            End If
        End If
    End Sub
    Sub bloquearInicio()
        DataGridView1.ReadOnly = True
        DataGridView2.ReadOnly = True
        Button1.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        TextBox2.Enabled = False
    End Sub
    Sub bloquear()
        DataGridView1.ReadOnly = True
        DataGridView2.ReadOnly = True
        Button1.Enabled = False
        Button2.Enabled = False
        Button4.Enabled = False
        TextBox1.Enabled = False
    End Sub
    Sub habilitar()
        DataGridView1.ReadOnly = False
        DataGridView2.ReadOnly = False
        Button1.Enabled = True
        Button2.Enabled = True
        Button4.Enabled = True
        TextBox2.Enabled = True
        Button6.Enabled = True
    End Sub
    Sub limpiar()
        DataGridView1.DataSource = Nothing
        DataGridView2.DataSource = Nothing
        TextBox1.Enabled = True
        TextBox2.Text = ""
        TextBox1.Text = ""
    End Sub
    Sub mostrarInformacionDetalle()
        da2.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select op_mca as OP,ISNULL( CONVERT(CHAR(70) , F.color_16B),'TODOS') AS COLOR,ISNULL( CONVERT(CHAR(70) , G.Talla_16C),'TODOS') AS TALLA,citem_mca as IT,linea_mca AS CODIGO,c.nomb_17 AS DESCRIPCION,factor_mca AS FACTOR,prendas_mca AS PRENDAS,cantreq_mca AS REQUERIDO,unid_mca AS UND, po_mca, isnull( b.ccolor_mca,'99'),isnull( b.ctalla_mca,'99')  from custom_vianny.dbo.requerimiento_logistico_analitico b
                                     LEFT JOIN custom_vianny.dbo.cag1700 C 
                                     ON b.ccia_mca = C.ccia AND b.linea_mca = C.linea_17
                                     LEFT join custom_vianny.dbo.qag160c G
                                     ON G.ccia + G.ncom_16c =  b.ccia_mca+ b.op_mca and b.ctalla_mca= G.correl_16c
                                     LEFT join custom_vianny.dbo.qag160b F 
                                     ON  F.ccia + F.ncom_16b = b.ccia_mca+ b.op_mca and F.CORREL_16B=b.ccolor_mca
                                        where ccia_mca ='" + Label3.Text.ToString().Trim() + "' and po_mca ='" + TextBox1.Text.ToString().Trim() + "' ORDER BY op_mca,citem_mca", conx)
        cmd.Fill(da2)
        If da2.Rows.Count > 0 Then
            DataGridView1.DataSource = da2
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).Width = 70
            DataGridView1.Columns(2).Width = 70
            DataGridView1.Columns(3).Width = 30
            DataGridView1.Columns(4).Width = 160
            DataGridView1.Columns(5).Width = 320
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(8).Width = 80
            DataGridView1.Columns(9).Width = 40
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(8).DefaultCellStyle.BackColor = Color.DarkSlateBlue
            DataGridView1.Columns(8).DefaultCellStyle.ForeColor = Color.White
        Else
            DataGridView1.DataSource = Nothing
        End If


    End Sub
    Sub mostrarInformacionTotal()
        da3.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select linea_mcr as CODIGO,c.nomb_17 AS Prodcuto,UniMedO_mcr AS Um,prendas_mcr AS Prendas,cantreq_mcr as Requerido,factconv_mcr AS 'Fconv. (Yardas)',cantreqo_mcr AS 'Req. Convertido',cantadic_mcr AS Adicional,totallog_mcr AS 'Total Comprar', po_mcr,UniMedO_mcr AS 'Um.' from custom_vianny.dbo.requerimiento_logistico_resumen r
                                        left join custom_vianny.dbo.cag1700  c
                                        on  r.ccia_mcr = c.ccia and r.linea_mcr = c.linea_17
                                        where ccia_mcr ='" + Label3.Text.ToString().Trim() + "' and po_mcr ='" + TextBox1.Text.ToString().Trim() + "'", conx)
        cmd.Fill(da3)
        If da3.Rows.Count > 0 Then
            DataGridView2.DataSource = da3
            DataGridView2.Columns(0).Width = 140
            DataGridView2.Columns(1).Width = 300
            DataGridView2.Columns(2).Width = 45
            DataGridView2.Columns(3).Width = 75
            DataGridView2.Columns(4).Width = 75
            DataGridView2.Columns(5).Width = 75
            DataGridView2.Columns(6).Width = 75
            DataGridView2.Columns(7).Width = 75
            DataGridView2.Columns(8).Width = 75
            DataGridView2.Columns(10).Width = 45
            DataGridView2.Columns(9).Visible = False
            DataGridView2.Columns(0).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(0).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(1).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(1).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(2).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(2).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(3).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(3).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(4).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(4).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(6).DefaultCellStyle.BackColor = Color.Beige
            DataGridView2.Columns(6).DefaultCellStyle.ForeColor = Color.Indigo
            DataGridView2.Columns(8).DefaultCellStyle.BackColor = Color.DarkSlateBlue
            DataGridView2.Columns(8).DefaultCellStyle.ForeColor = Color.White

        Else
            DataGridView2.DataSource = Nothing
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscarOP()
    End Sub

    Sub mostrarDetallado()
        da1.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("exec ExplosionAviosOp  '" + TextBox1.Text.ToString().Trim() + "','" + Label3.Text.ToString().Trim() + "'", conx)
        cmd.Fill(da1)
        If da1.Rows.Count > 0 Then
            DataGridView1.DataSource = da1
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).Width = 70
            DataGridView1.Columns(2).Width = 70
            DataGridView1.Columns(3).Width = 30
            DataGridView1.Columns(4).Width = 160
            DataGridView1.Columns(5).Width = 320
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(8).Width = 80
            DataGridView1.Columns(9).Width = 40
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(8).DefaultCellStyle.BackColor = Color.DarkSlateBlue
            DataGridView1.Columns(8).DefaultCellStyle.ForeColor = Color.White
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Function ELIMINAR(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        abrir()
        Dim agregar As String = "delete FROM custom_vianny.dbo.Requerimiento_Logistico_Resumen Where CCia_mcr='" + Label3.Text.ToString().Trim() + "' AND po_mcr='" + TextBox1.Text.ToString().Trim() + "' "
        For i = 0 To DataGridView2.Rows.Count - 1
            Dim cmd15 As New SqlCommand("insert into custom_vianny.dbo.Requerimiento_Logistico_Resumen (ccia_mcr,po_mcr,linea_mcr,prendas_mcr,cantreq_mcr,factconv_mcr,cantreqo_mcr,UniMedO_mcr,cantadic_mcr,totallog_mcr)
                                    values (@ccia_mcr,@po_mcr,@linea_mcr,@prendas_mcr,@cantreq_mcr,@factconv_mcr,@cantreqo_mcr,@UniMedO_mcr,@cantadic_mcr,@totallog_mcr)", conx)
            cmd15.Parameters.AddWithValue("@ccia_mcr", Label3.Text.ToString().Trim())
            cmd15.Parameters.AddWithValue("@po_mcr", DataGridView1.Rows(i).Cells(7).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@linea_mcr", DataGridView1.Rows(i).Cells(0).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@prendas_mcr", Convert.ToDecimal(DataGridView1.Rows(i).Cells(2).Value))
            cmd15.Parameters.AddWithValue("@cantreq_mcr", Convert.ToDecimal(DataGridView1.Rows(i).Cells(3).Value))
            cmd15.Parameters.AddWithValue("@factconv_mcr", Convert.ToDecimal(DataGridView1.Rows(i).Cells(5).Value))
            cmd15.Parameters.AddWithValue("@cantreqo_mcr", Convert.ToDecimal(DataGridView1.Rows(i).Cells(6).Value))
            cmd15.Parameters.AddWithValue("@UniMedO_mcr", DataGridView1.Rows(i).Cells(4).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@cantadic_mcr", 0)
            cmd15.Parameters.AddWithValue("@totallog_mcr", Convert.ToDecimal(DataGridView1.Rows(i).Cells(6).Value))
            cmd15.ExecuteNonQuery()
        Next
        For i1 = 0 To DataGridView2.Rows.Count - 1
            Dim cmd16 As New SqlCommand("insert into custom_vianny.dbo.Requerimiento_Logistico_Analitico (ccia_mca,po_mca,citem_mca,op_mca,linea_mca,unid_mca,nodecla_mca,ctalla_mca,ccolor_mca,factor_mca,prendas_mca,cantreq_mca)
										values (@ccia_mca,@po_mca,@citem_mca,@op_mca,@linea_mca,@unid_mca,@nodecla_mca,@ctalla_mca,@ccolor_mca,@factor_mca,@prendas_mca,@cantreq_mca)", conx)
            cmd16.Parameters.AddWithValue("@ccia_mca", Label3.Text.ToString().Trim())
            cmd16.Parameters.AddWithValue("@po_mca", DataGridView1.Rows(i1).Cells(10).Value.ToString().Trim())
            cmd16.Parameters.AddWithValue("@citem_mca", DataGridView1.Rows(i1).Cells(3).Value.ToString().Trim())
            cmd16.Parameters.AddWithValue("@op_mca", DataGridView1.Rows(i1).Cells(0).Value.ToString().Trim())
            cmd16.Parameters.AddWithValue("@linea_mca", DataGridView1.Rows(i1).Cells(4).Value.ToString().Trim())
            cmd16.Parameters.AddWithValue("@unid_mca", DataGridView1.Rows(i1).Cells(9).Value.ToString().Trim())
            cmd16.Parameters.AddWithValue("@nodecla_mca", 0)
            cmd16.Parameters.AddWithValue("@ctalla_mca", DataGridView1.Rows(i1).Cells(12).Value.ToString().Trim())
            cmd16.Parameters.AddWithValue("@ccolor_mca", DataGridView1.Rows(i1).Cells(11).Value.ToString().Trim())
            cmd16.Parameters.AddWithValue("@factor_mca", Convert.ToDecimal(DataGridView1.Rows(i1).Cells(6).Value))
            cmd16.Parameters.AddWithValue("@prendas_mca", Convert.ToDecimal(DataGridView1.Rows(i1).Cells(7).Value))
            cmd16.Parameters.AddWithValue("@cantreq_mca", Convert.ToDecimal(DataGridView1.Rows(i1).Cells(8).Value))
            cmd16.ExecuteNonQuery()
        Next
        MsgBox("SE GUARDO LA INFORMACION CORRECTAMENTE")
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        TextBox1.Text = ""
        TextBox1.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
    End Sub

    Public Sub ExportarDatosPDF(ByVal document As Document)
        'Se crea un objeto PDFTable con el numero de columnas del DataGridView. 

        Dim fuentePredeterminada As New Font(Font.Name = "Tahoma", 8, Font.Bold)
        Dim datatable As New PdfPTable(DataGridView2.ColumnCount)
        'Se asignan algunas propiedades para el diseño del PDF.
        datatable.DefaultCell.Padding = 3
        Dim headerwidths As Single() = GetColumnasSize(DataGridView2) ' Aqui se realiza una llamada a la funcion GetColumnasSize y le pasamos como parametro nuestro datagridview
        datatable.SetWidths(headerwidths) 'Pasamos como parametro el array que contiene los ancho de columna, a la propiedad "SetWidths"
        datatable.WidthPercentage = 100
        datatable.DefaultCell.BorderWidth = 0.5
        datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER

        'Se crea el encabezado en el PDF. 
        Dim encabezado As New Paragraph("CONSORCIO TEXTIL VIANNY S.A.C", New Font(Font.Name = "Tahoma", 14, Font.Bold))

        'Se crea el texto abajo del encabezado.
        Dim texto As New Phrase("Explision de Avios :" + "Po :  " + TextBox1.Text, New Font(Font.Name = "Tahoma", 9, Font.Bold))

        'Se capturan los nombres de las columnas del DataGridView.
        'For i As Integer = 0 To DataGridView2.ColumnCount - 1
        '    datatable.AddCell(DataGridView2.Columns(i).HeaderText)
        'Next
        For i As Integer = 0 To 8
            Dim cell As New PdfPCell(New Phrase(DataGridView2.Columns(i).HeaderText, fuentePredeterminada))
            cell.HorizontalAlignment = Element.ALIGN_CENTER
            cell.BackgroundColor = BaseColor.LIGHT_GRAY
            datatable.AddCell(cell)
        Next
        datatable.HeaderRows = 1
        datatable.DefaultCell.BorderWidth = 0.1

        'Se generan las columnas del DataGridView. 
        For i As Integer = 0 To DataGridView2.RowCount - 1
            For j As Integer = 0 To 8
                Dim valorCelda As String = If(DataGridView2(j, i).Value IsNot Nothing, DataGridView2(j, i).Value.ToString(), "")
                Dim cell As New PdfPCell(New Phrase(valorCelda, fuentePredeterminada))
                cell.HorizontalAlignment = Element.ALIGN_CENTER
                datatable.AddCell(cell)
            Next
            datatable.CompleteRow()
        Next

        'Se agrega el PDFTable al documento.
        document.Add(encabezado)
        document.Add(texto)
        document.Add(datatable)
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        habilitar()
    End Sub

    Private Sub ExplosionAvios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bloquearInicio()
    End Sub
    Public Function GetColumnasSize(ByVal dg As DataGridView) As Single() 'Se obtiene el ancho de las columnas del Datagridview
        Dim values As Single() = New Single(dg.ColumnCount - 1) {} 'Se declara un array vacio de tipo "Single" cuyo tamaño sera el numero de columnas del datagridview
        For i As Integer = 0 To 8 'Con un ciclo for recorremos las columnas del datagridview
            values(i) = CSng(dg.Columns(i).Width) 'Se obtiene y se convierte el ancho de la columna a tipo "Single" para cargarlo en el array "values"
        Next
        Return values 'Y por ultimo nos retorna un array, que contiene el ancho de cada columna del datagridview.
    End Function
    Private Sub AltoFila() 'Para configurar el alto de la fila del datagridview
        For Each row As DataGridViewRow In DataGridView1.Rows
            'Se asigna el alto de la fila para poder ver la imagen correctamente
            row.Height = 120
        Next
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            'Intentar generar el documento.
            Dim doc As New Document(PageSize.A4.Rotate(), 10, 10, 10, 10)

            'Path que guarda el reporte en el escritorio de windows (Desktop).
            Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\Reporteproductos.pdf"

            Dim file As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
            PdfWriter.GetInstance(doc, file)
            doc.Open()
            ExportarDatosPDF(doc)
            doc.Close()
            Process.Start(filename)

        Catch ex As Exception
            'Si el intento es fallido, mostrar MsgBox.
            MessageBox.Show("No se puede generar el documento PDF.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellEndEdit
        If DataGridView2.Columns(5).HeaderText = "Fconv. (Yardas)" Then
            If DataGridView2.Rows(e.RowIndex).Cells(5).Value IsNot Nothing AndAlso DataGridView2.Rows(e.RowIndex).Cells(5).Value.ToString().Trim().Length > 0 Then
                Dim valor As Double
                valor = Convert.ToDouble(DataGridView2.Rows(e.RowIndex).Cells(4).Value) / Convert.ToDouble(DataGridView2.Rows(e.RowIndex).Cells(5).Value)
                DataGridView2.Rows(e.RowIndex).Cells(6).Value = Math.Ceiling(valor).ToString()
                DataGridView2.Rows(e.RowIndex).Cells(8).Value = Math.Ceiling(valor).ToString()
                DataGridView2.Rows(e.RowIndex).Cells(10).Value = "CON"
            Else
                DataGridView2.Rows(e.RowIndex).Cells(6).Value = DataGridView2.Rows(e.RowIndex).Cells(4).Value
                DataGridView2.Rows(e.RowIndex).Cells(8).Value = DataGridView2.Rows(e.RowIndex).Cells(4).Value
                DataGridView2.Rows(e.RowIndex).Cells(10).Value = DataGridView2.Rows(e.RowIndex).Cells(2).Value
            End If

        End If
        If DataGridView2.Columns(7).HeaderText = "Adicional" Then
            If DataGridView2.Rows(e.RowIndex).Cells(7).Value IsNot Nothing AndAlso DataGridView2.Rows(e.RowIndex).Cells(7).Value.ToString().Trim().Length > 0 Then
                Dim valor1 As Double
                valor1 = Convert.ToDouble(DataGridView2.Rows(e.RowIndex).Cells(6).Value) + Convert.ToDouble(DataGridView2.Rows(e.RowIndex).Cells(7).Value)
                DataGridView2.Rows(e.RowIndex).Cells(8).Value = Math.Ceiling(valor1).ToString()
            End If

        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        limpiar()

    End Sub

    Private Sub buscarOP()
        Try
            Dim ds As New DataSet
            Dim K As Integer
            ds.Tables.Add(da1.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "OP" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 80
                DataGridView1.Columns(1).Width = 70
                DataGridView1.Columns(2).Width = 70
                DataGridView1.Columns(3).Width = 30
                DataGridView1.Columns(4).Width = 160
                DataGridView1.Columns(5).Width = 320
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(8).Width = 80
                DataGridView1.Columns(9).Width = 40
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(8).DefaultCellStyle.BackColor = Color.DarkSlateBlue
                DataGridView1.Columns(8).DefaultCellStyle.ForeColor = Color.White
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
End Class