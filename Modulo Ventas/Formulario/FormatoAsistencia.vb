Imports System.Text
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Public Class FormatoAsistencia
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Trim(TextBox2.Text).Length = 0 And Trim(TextBox3.Text).Length = 0 Then
            MsgBox("NO HAY DATOS EN DNI NI TRABAJADOR")
        Else
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()

            ' Obtener el año y mes seleccionados (o usar el año y mes actuales)
            Dim year As Integer = TextBox1.Text
            Dim month As Integer
            Select Case ComboBox1.Text
                Case "ENERO" : month = "01"
                Case "FEBRERO" : month = "02"
                Case "MARZO" : month = "03"
                Case "ABRIL" : month = "04"
                Case "MAYO" : month = "05"
                Case "JUNIO" : month = "06"
                Case "JULIO" : month = "07"
                Case "AGOSTO" : month = "08"
                Case "SETIEMBRE" : month = "09"
                Case "OCTUBRE" : month = "10"
                Case "NOVIEMBRE" : month = "11"
                Case "DICIEMBRE" : month = "12"
            End Select


            ' Obtener el número de días en el mes
            Dim daysInMonth As Integer = DateTime.DaysInMonth(year, month)

            ' Agregar una columna al DataGridView
            DataGridView1.Columns.Add("Fecha", "Fecha")
            DataGridView1.Columns.Add("Dni", "Dni")
            DataGridView1.Columns.Add("Trabajador", "Trabajador")
            DataGridView1.Columns.Add("Horario Ingreso", "Horario Ingreso")
            DataGridView1.Columns.Add("Horario Salida", "Horario Salida")
            'DataGridView1.Columns.Add("Ingreso Real", "Ingreso Real")
            'DataGridView1.Columns.Add("Sabado", "Sabado")
            ' Obtener el rango de horas y minutos seleccionados
            Dim startTime As DateTime = DateTimePicker1.Value
            Dim endTime As DateTime = DateTimePicker2.Value
            Dim startTime2 As DateTime = DateTimePicker6.Value
            Dim endTime2 As DateTime = DateTimePicker5.Value
            Dim startTime3 As DateTime = DateTimePicker3.Value
            Dim endTime3 As DateTime = DateTimePicker4.Value
            Dim rand As New Random()
            ' Asegurarse de que el rango de tiempo es válido
            If startTime > endTime Then
                MessageBox.Show("El tiempo de inicio debe ser menor que el tiempo de fin.")
                Return
            End If
            ' Llenar el DataGridView con los días del mes
            For day As Integer = 1 To daysInMonth
                Dim dateValue As New DateTime(year, month, day)
                Dim randomTime As DateTime = GetRandomTime(startTime, endTime, rand)
                Dim randomTime2 As DateTime = GetRandomTime2(startTime2, endTime2, rand)
                Dim randomTime3 As DateTime = GetRandomTime3(startTime3, endTime3, rand)
                If dateValue.DayOfWeek <> DayOfWeek.Sunday Then
                    Dim esSabado As String = If(dateValue.DayOfWeek = DayOfWeek.Saturday, "Sí", "No")
                    If esSabado = "Sí" Then
                        DataGridView1.Rows.Add(dateValue.ToShortDateString(), Trim(TextBox2.Text), Trim(TextBox3.Text), randomTime.ToString("HH:mm"), randomTime3.ToString("HH:mm"), esSabado)

                    Else
                        DataGridView1.Rows.Add(dateValue.ToShortDateString(), Trim(TextBox2.Text), Trim(TextBox3.Text), randomTime.ToString("HH:mm"), randomTime2.ToString("HH:mm"), esSabado)

                    End If
                End If
            Next
            DataGridView1.Columns(2).Width = 250
        End If


        ' Generar horarios aleatorios

        'For i As Integer = 1 To 10 ' Generar 10 horarios aleatorios

        'Next
    End Sub
    Private Function GetRandomTime(startTime As DateTime, endTime As DateTime, rand As Random) As DateTime
        Dim totalMinutes As Integer = CInt((endTime - startTime).TotalMinutes)
        Dim randomMinutes As Integer = rand.Next(0, totalMinutes)
        Return startTime.AddMinutes(randomMinutes)
    End Function
    Private Function GetRandomTime2(startTime2 As DateTime, endTime2 As DateTime, rand2 As Random) As DateTime
        Dim totalMinutes2 As Integer = CInt((endTime2 - startTime2).TotalMinutes)
        Dim randomMinutes2 As Integer = rand2.Next(0, totalMinutes2)
        Return startTime2.AddMinutes(randomMinutes2)
    End Function
    Private Function GetRandomTime3(startTime3 As DateTime, endTime3 As DateTime, rand3 As Random) As DateTime
        Dim totalMinutes3 As Integer = CInt((endTime3 - startTime3).TotalMinutes)
        Dim randomMinutes3 As Integer = rand3.Next(0, totalMinutes3)
        Return startTime3.AddMinutes(randomMinutes3)
    End Function

    Private Sub FormatoAsistencia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DateTimePicker1.Format = DateTimePickerFormat.Time
        DateTimePicker1.ShowUpDown = True
        DateTimePicker2.Format = DateTimePickerFormat.Time
        DateTimePicker2.ShowUpDown = True
        DateTimePicker6.Format = DateTimePickerFormat.Time
        DateTimePicker6.ShowUpDown = True
        DateTimePicker5.Format = DateTimePickerFormat.Time
        DateTimePicker5.ShowUpDown = True
        DateTimePicker3.Format = DateTimePickerFormat.Time
        DateTimePicker3.ShowUpDown = True
        DateTimePicker4.Format = DateTimePickerFormat.Time
        DateTimePicker4.ShowUpDown = True
    End Sub
    Public Sub ExportarDatosPDF(ByVal document As Document)
        'Se crea un objeto PDFTable con el numero de columnas del DataGridView. 
        Dim datatable As New PdfPTable(DataGridView1.ColumnCount)
        'Se asignan algunas propiedades para el diseño del PDF.
        datatable.DefaultCell.Padding = 3
        Dim headerwidths As Single() = GetColumnasSize(DataGridView1) ' Aqui se realiza una llamada a la funcion GetColumnasSize y le pasamos como parametro nuestro datagridview
        datatable.SetWidths(headerwidths) 'Pasamos como parametro el array que contiene los ancho de columna, a la propiedad "SetWidths"
        datatable.WidthPercentage = 100
        datatable.DefaultCell.BorderWidth = 2
        datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER
        'Se crea el encabezado en el PDF. 
        Dim encabezado As New Paragraph("CONSORCIO TEXTIL VIANNY S.A.C", New Font(Font.Name = "Tahoma", 20, Font.Bold))

        'Se crea el texto abajo del encabezado.
        Dim texto As New Phrase("lISTA DE MARCACIONES :" + "PERIODO :  " + TextBox1.Text + "  MES :  " + ComboBox1.Text, New Font(Font.Name = "Tahoma", 14, Font.Bold))

        'Se capturan los nombres de las columnas del DataGridView.
        For i As Integer = 0 To DataGridView1.ColumnCount - 1
            datatable.AddCell(DataGridView1.Columns(i).HeaderText)
        Next

        datatable.HeaderRows = 1
        datatable.DefaultCell.BorderWidth = 1

        'Se generan las columnas del DataGridView. 
        For i As Integer = 0 To DataGridView1.RowCount - 1 'Recorre la filas del datagridview
            For j As Integer = 0 To DataGridView1.ColumnCount - 1 'Recorre las columnas del datagridview

                'compruebo que columna contien la imagen y en que columna deseo que se muestre
                'If (j = 4) Then
                '    'capturo la ruta de la imagen
                '    Dim RutaImage As String
                '    RutaImage = DataGridView1("Column5", i).Value
                '    'Procedo a convertir a imagen de tipo itextsharp.text.image
                '    Dim Img As Image = Image.GetInstance(RutaImage)
                '    'agrego la imagen a la celda
                '    datatable.AddCell(Img)
                'Else
                datatable.AddCell(DataGridView1(j, i).Value.ToString())
                'End If

            Next
            datatable.CompleteRow()
        Next
        'Se agrega el PDFTable al documento.
        document.Add(encabezado)
        document.Add(texto)
        document.Add(datatable)
    End Sub
    Public Function GetColumnasSize(ByVal dg As DataGridView) As Single() 'Se obtiene el ancho de las columnas del Datagridview
        Dim values As Single() = New Single(dg.ColumnCount - 1) {} 'Se declara un array vacio de tipo "Single" cuyo tamaño sera el numero de columnas del datagridview
        For i As Integer = 0 To dg.ColumnCount - 1 'Con un ciclo for recorremos las columnas del datagridview
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

  
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
        Catch ex As Exception
            'Si el intento es fallido, mostrar MsgBox.
            MessageBox.Show("No se puede generar el documento PDF.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class