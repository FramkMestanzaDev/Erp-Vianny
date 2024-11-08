Public Class ExportaruDP
    Function llenarExcel(ByVal ElGrid As DataGridView) As Boolean

        'Creamos las variables

        Dim exApp As New Microsoft.Office.Interop.Excel.Application

        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook

        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        Try

            'Añadimos el Libro al programa, y la hoja al libro

            exLibro = exApp.Workbooks.Add

            exHoja = exLibro.Worksheets.Add()

            ' ¿Cuantas columnas y cuantas filas?

            Dim NCol As Integer = ElGrid.ColumnCount

            Dim NRow As Integer = ElGrid.RowCount

            'Aqui recorremos todas las filas, y por cada fila todas las columnas

            ' y vamos escribiendo.
            exHoja.Columns("K:K").NumberFormat = "DD/mm/yyyy;@"
            exHoja.Columns("L:L").NumberFormat = "DD/mm/yyyy;@"
            exHoja.Columns("P:P").NumberFormat = "DD/mm/yyyy;@"
            exHoja.Columns("Q:Q").NumberFormat = "DD/mm/yyyy;@"
            For i As Integer = 1 To NCol

                exHoja.Cells.Item(1, i) = ElGrid.Columns(i - 1).Name.ToString

            Next

            For Fila As Integer = 0 To NRow - 1

                For Col As Integer = 0 To NCol - 1
                    If Col = 10 Or Col = 15 Or Col = 16 Or Col = 11 Then
                        exHoja.Cells.Item(Fila + 2, Col + 1) = "'" + Trim(ElGrid.Item(Col, Fila).Value)
                    Else
                        exHoja.Cells.Item(Fila + 2, Col + 1) = Trim(ElGrid.Item(Col, Fila).Value)
                    End If


                Next

            Next

            'Titulo en negrita, Alineado al centro y que el tamaño de la columna

            'se ajuste al texto

            exHoja.Rows.Item(1).Font.Bold = 1

            exHoja.Rows.Item(1).HorizontalAlignment = 3

            exHoja.Columns.AutoFit()

            ' Aplicación visible

            exApp.Application.Visible = True

            exHoja = Nothing

            exLibro = Nothing

            exApp = Nothing

        Catch ex As Exception

            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False

        End Try

        Return True

    End Function
End Class
