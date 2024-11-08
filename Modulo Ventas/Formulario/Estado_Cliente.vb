Public Class Estado_Cliente
    Dim dt, dt1 As New DataTable
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = 7
            Clientes.Show()
        Else
            If e.KeyCode = Keys.Enter Then
                Dim jk As New fcliente
                Dim hh As New vcliente
                hh.gruc = TextBox4.Text
                TextBox1.Text = jk.BUSCAR_cliente_gfff(hh)
            End If
        End If
    End Sub

    Private Sub Estado_Cliente_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox4.Select()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("EL CAMPO CLIENT ESTA VACIO")
        Else
            Dim fk As New festadocuenta
            Dim fk1 As New vestadocuenta
            Dim AD As Integer
            Dim SUMA, suma1, suma2 As Double
            fk1.gruc = TextBox4.Text
            fk1.gFECHA = DateTimePicker1.Value
            fk1.gccia = Label5.Text


            dt = fk.buscar_estado_cuenta(fk1)
            If dt Is Nothing Then
                MsgBox("NO EXISTE VENTAS  POR COBRAR")
            Else

                If dt.Rows.Count <> 0 Then
                    DataGridView1.DataSource = dt
                    DataGridView1.Columns(1).Width = 190
                    DataGridView1.Columns(3).Width = 220
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(2).Visible = False
                    DataGridView1.Columns(4).Visible = False
                    DataGridView1.Columns(6).Visible = False
                    DataGridView1.Columns(8).Visible = False
                    DataGridView1.Columns(9).Visible = False
                    DataGridView1.Columns(14).Visible = False
                    DataGridView1.Columns(15).Visible = False
                    DataGridView1.Columns(16).Visible = False
                    DataGridView1.Columns(17).Visible = False
                    DataGridView1.Columns(20).Visible = False
                    DataGridView1.Columns(21).Visible = False
                    DataGridView1.Columns(22).Visible = False
                    DataGridView1.Columns(23).Visible = False
                    DataGridView1.Columns(24).Visible = False
                    AD = DataGridView1.Rows.Count
                    For i = 0 To AD - 1
                        SUMA = SUMA + DataGridView1.Rows(i).Cells(18).Value
                        suma1 = suma1 + DataGridView1.Rows(i).Cells(19).Value
                        suma2 = suma2 + DataGridView1.Rows(i).Cells(6).Value
                    Next

                    TextBox6.Text = Format(SUMA, "##,##00.00")

                    TextBox2.Text = Format(suma1, "##,##00.00")
                Else
                    MsgBox("NO EXISTE VENTAS  POR COBRAR")
                End If
            End If
        End If


    End Sub
    Public Function GetColumnasSize(ByVal dg As DataGridView) As Single()
        Dim values As Single() = New Single(dg.ColumnCount - 1) {}
        For i As Integer = 0 To dg.ColumnCount - 1
            values(i) = CSng(dg.Columns(i).Width)
        Next
        Return values
    End Function
    'Public Sub ExportarDatosPDF(ByVal document As Document)
    '    'Se crea un objeto PDFTable con el numero de columnas del DataGridView. 

    '    Dim datatable As New PdfPTable(DataGridView1.ColumnCount)
    '    'Se asignan algunas propiedades para el diseño del PDF.
    '    datatable.DefaultCell.Padding = 3
    '    Dim headerwidths As Single() = GetColumnasSize(DataGridView1)
    '    datatable.SetWidths(headerwidths)
    '    datatable.WidthPercentage = 90
    '    datatable.DefaultCell.BorderWidth = 1
    '    datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER

    '    'datatable.SetWidths(New Single() {7.0F, 5.5F, 9.5F})


    '    'datatable.TotalWidth = 500.0F
    '    'datatable.LockedWidth = True
    '    datatable.SpacingBefore = 7.0F
    '    datatable.HorizontalAlignment = Element.ALIGN_LEFT
    '    'Se crea el encabezado en el PDF. 
    '    Dim encabezado As New Paragraph("ESTADO DE CUENTA DEL CLIENTE", New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    encabezado.Alignment = Element.ALIGN_CENTER

    '    'Se crea el texto abajo del encabezado.
    '    Dim texto1 As New Paragraph("CONSORCIO TEXTIL VIANNY", New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    Dim texto As New Phrase("Fecha Reporte :" + Now.Date(), New Font(Font.Name = "Tahoma", 9, Font.Bold))
    '    Dim texto2 As New Paragraph("VENTAS CON FACTURA", New Font(Font.Name = "Tahoma", 13, Font.Bold))
    '    texto2.Alignment = Element.ALIGN_CENTER

    '    'Se capturan los nombres de las columnas del DataGridView.
    '    For i As Integer = 0 To DataGridView1.ColumnCount - 1
    '        datatable.AddCell(DataGridView1.Columns(i).HeaderText)
    '        'Dim celda As New PdfPCell(New Phrase(DataGridView1.Columns(i).HeaderText, New Font(Font.Name = "HELVETICA", 12, Font.Bold)))

    '        'datatable.AddCell(celda)
    '    Next
    '    datatable.HeaderRows = 1
    '    datatable.DefaultCell.BorderWidth = 0


    '    'Se generan las columnas del DataGridView. 
    '    For i As Integer = 0 To DataGridView1.RowCount - 1
    '        For j As Integer = 0 To DataGridView1.ColumnCount - 1
    '            'compruebo que columna contien la imagen y en que columna deseo que se muestre

    '            'datatable.AddCell(DataGridView1(j, i).Value.ToString())
    '            Dim celda As New PdfPCell(New Phrase(DataGridView1(j, i).Value.ToString(), New Font(Font.Name = "Tahoma", 9)))
    '            datatable.AddCell(celda)
    '        Next

    '    Next

    '    'Se agrega el PDFTable al documento.
    '    document.Add(encabezado)
    '    document.Add(texto1)
    '    document.Add(texto)
    '    document.Add(texto2)
    '    document.Add(datatable)
    'End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim jk As New Exportar
        jk.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Form_cuenta_cliente.TextBox1.Text = DateTimePicker1.Value.Date.ToString("yyyy/MM/dd")

        Form_cuenta_cliente.TextBox2.Text = Trim(TextBox4.Text)
        Form_cuenta_cliente.Show()
        'Try
        '    'Intentar generar el documento.
        '    Dim IA As Integer
        '    Dim fuente As iTextSharp.text.pdf.BaseFont
        '    Dim doc As New Document(PageSize.A4.Rotate(), 10, 10, 10, 10)
        '    'Path que guarda el reporte en el escritorio de windows (Desktop).
        '    Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\Reporteproductos.pdf"
        '    Dim file As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
        '    Dim cb As PdfContentByte

        '    PdfWriter.GetInstance(doc, file)
        '    doc.Open()

        '    fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
        '    'Seteamos el tipo de letra y el tamaño.

        '    ExportarDatosPDF(doc)
        '    IA = DataGridView1.Rows.Count
        '    If IA > 0 Then
        '        ExportarDatosPDF(doc)

        '    End If

        '    'ExportarDatosPDF2(doc)
        '    'ExportarDatosPDF3(doc)
        '    'ExportarDatosPDF4(doc)

        '    doc.Close()
        '    Process.Start(filename)
        'Catch ex As Exception
        '    'Si el intento es fallido, mostrar MsgBox.
        '    MessageBox.Show("No se puede generar el documento PDF.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
    End Sub

    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        ToolTip1.SetToolTip(PictureBox1, "EXPORTAR EXCEL")
        ToolTip1.ToolTipTitle = "ESTADO CUENTA"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Form_cuenta_cliente.TextBox1.Text = DateTimePicker1.Value.Date.ToString("yyyy/MM/dd")
        Form_cuenta_cliente.TextBox2.Text = Trim(TextBox4.Text)
        Form_cuenta_cliente.TextBox3.Text = Label5.Text
        Form_cuenta_cliente.Show()
    End Sub

    Private Sub Button2_MouseHover(sender As Object, e As EventArgs) Handles Button2.MouseHover
        ToolTip1.SetToolTip(PictureBox1, "IMPRIMIR PDF")
        ToolTip1.ToolTipTitle = "ESTADO CUENTA"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub
End Class