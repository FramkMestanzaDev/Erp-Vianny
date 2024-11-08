Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Public Class Morosidad_Clientes
    'Dim dt, dt1 As New DataTable

    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    Dim hu As New fventas
    '    Dim hg As New vventas
    '    Dim t As New Integer

    '    Select Case ComboBox1.Text.ToString.Trim
    '        Case "VPIZARRO" : hg.gVENDEDOR = "0027"
    '        Case "VSILVERIO" : hg.gVENDEDOR = "0005"
    '        Case "GBALVIN" : hg.gVENDEDOR = "0028"
    '        Case "GBEDON" : hg.gVENDEDOR = "0010"
    '        Case "VINCIO" : hg.gVENDEDOR = "0022"
    '        Case "DBRAVO" : hg.gVENDEDOR = "0023"
    '        Case "AMENDO" : hg.gVENDEDOR = "0026"
    '        Case "GCUEVA" : hg.gVENDEDOR = "0029"
    '        Case "JSALINAS" : hg.gVENDEDOR = "0025"
    '        Case "VGRAUS" : hg.gVENDEDOR = "0007"
    '        Case "RMEDINA" : hg.gVENDEDOR = "0030"
    '    End Select
    '    dt = hu.buscar_ventas_cliente(hg)

    '    DataGridView1.DataSource = dt
    '    DataGridView1.Columns(1).DefaultCellStyle.Format = "#,##0.00"
    '    DataGridView1.Columns(2).DefaultCellStyle.Format = "#,##0.00"
    '    DataGridView1.Columns(3).DefaultCellStyle.Format = "#,##0.00"
    '    DataGridView1.Columns(3).DefaultCellStyle.BackColor = Color.LightBlue
    '    DataGridView1.Columns(3).DefaultCellStyle.ForeColor = Color.Blue
    '    DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    '    DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    '    DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    '    DataGridView1.Columns(0).Width = 450
    '    DataGridView1.Columns(1).Width = 110
    '    DataGridView1.Columns(2).Width = 110
    '    DataGridView1.Columns(3).Width = 110
    '    'Me.Chart1.Series.Clear()
    '    'Me.Chart1.ChartAreas.Clear()
    '    'Me.Chart1.Series.Add("1")
    '    'Chart1.Series.Add("2")
    '    'Dim area As New ChartArea()
    '    'Chart1.ChartAreas().Add(area)
    '    'AD1 = Convert.ToInt32(DataGridView3.ColumnCount.ToString())
    '    'For i = 0 To AD1 - 1
    '    '    Me.Chart1.Series("1").Points.AddXY(i + 1, DataGridView3.Rows(0).Cells(i).Value)
    '    '    Me.Chart1.Series("2").Points.AddXY(i + 1, DataGridView3.Rows(1).Cells(i).Value)
    '    '    'SUMA20 = SUMA20 + DataGridView1.Rows(1).Cells(i).Value

    '    'Next

    'End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    Try
    '        'Intentar generar el documento.

    '        Dim fuente As iTextSharp.text.pdf.BaseFont
    '        Dim doc As New Document(PageSize.A4.Rotate(), 10, 10, 10, 10)
    '        'Path que guarda el reporte en el escritorio de windows (Desktop).
    '        Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\Reporteproductos.pdf"
    '        Dim file As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)


    '        PdfWriter.GetInstance(doc, file)
    '        doc.Open()

    '        fuente = FontFactory.GetFont(FontFactory.HELVETICA, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL).BaseFont
    '        'Seteamos el tipo de letra y el tamaño.

    '        ExportarDatosPDF1(doc)


    '        'ExportarDatosPDF2(doc)
    '        'ExportarDatosPDF3(doc)
    '        'ExportarDatosPDF4(doc)

    '        doc.Close()
    '        Process.Start(filename)
    '    Catch ex As Exception
    '        'Si el intento es fallido, mostrar MsgBox.
    '        MessageBox.Show("No se puede generar el documento PDF.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub
    'Public Function GetColumnasSize(ByVal dg As DataGridView) As Single()
    '    Dim values As Single() = New Single(dg.ColumnCount - 1) {}
    '    For i As Integer = 0 To dg.ColumnCount - 1
    '        values(i) = CSng(dg.Columns(i).Width)
    '    Next
    '    Return values
    'End Function

    'Private Sub Morosidad_Clientes_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    'End Sub

    ''Public Sub ExportarDatosPDF1(ByVal document As Document)
    ''    'Se crea un objeto PDFTable con el numero de columnas del DataGridView. 

    ''    Dim datatable As New PdfPTable(DataGridView1.ColumnCount)
    ''    'Se asignan algunas propiedades para el diseño del PDF.
    ''    datatable.DefaultCell.Padding = 3
    ''    Dim headerwidths As Single() = GetColumnasSize(DataGridView1)
    ''    datatable.SetWidths(headerwidths)
    ''    datatable.WidthPercentage = 90
    ''    datatable.DefaultCell.BorderWidth = 1
    ''    datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER

    ''    'datatable.SetWidths(New Single() {7.0F, 5.5F, 9.5F})


    ''    'datatable.TotalWidth = 500.0F
    ''    'datatable.LockedWidth = True
    ''    datatable.SpacingBefore = 7.0F
    ''    datatable.HorizontalAlignment = Element.ALIGN_LEFT
    ''    'Se crea 1el encabezado en el PDF. 
    ''    Dim texto1 As New Paragraph("MOROSIDAD DE CLIENTES", New Font(Font.Name = "Tahoma", 16, Font.Bold))
    ''    texto1.Alignment = Element.ALIGN_CENTER
    ''    Dim texto2 As New Paragraph("VENDEDOR : ---" + ComboBox1.Text, New Font(Font.Name = "Tahoma", 13, Font.Bold))
    ''    texto2.Alignment = Element.ALIGN_CENTER

    ''    'Se capturan los nombres de las columnas del DataGridView.
    ''    For i As Integer = 0 To DataGridView1.ColumnCount - 1
    ''        datatable.AddCell(DataGridView1.Columns(i).HeaderText)
    ''        'Dim celda As New PdfPCell(New Phrase(DataGridView1.Columns(i).HeaderText, New Font(Font.Name = "HELVETICA", 12, Font.Bold)))

    ''        'datatable.AddCell(celda)
    ''    Next
    ''    datatable.HeaderRows = 1
    ''    datatable.DefaultCell.BorderWidth = 0


    ''    'Se generan las columnas del DataGridView. 
    ''    For i As Integer = 0 To DataGridView1.RowCount - 1
    ''        For j As Integer = 0 To DataGridView1.ColumnCount - 1
    ''            'compruebo que columna contien la imagen y en que columna deseo que se muestre

    ''            'datatable.AddCell(DataGridView1(j, i).Value.ToString())
    ''            Dim celda As New PdfPCell(New Phrase(DataGridView1(j, i).Value.ToString(), New Font(Font.Name = "Tahoma", 9)))
    ''            datatable.AddCell(celda)
    ''        Next

    ''    Next

    ''    'Se agrega el PDFTable al documento.
    ''    document.Add(texto1)
    ''    document.Add(texto2)
    ''    document.Add(datatable)
    ''End Sub
End Class