Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Public Class Fecha__Factura
    Public conx As SqlConnection
    Dim Rsr24 As SqlDataReader
    Dim Rsr22, Rsr21 As SqlDataReader
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Configura el diálogo para abrir un archivo Excel
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "Archivos Excel|*.xls;*.xlsx|Todos los archivos|*.*"
        openFileDialog1.Title = "Seleccionar archivo Excel"

        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Obtiene la ruta del archivo seleccionado
            Dim archivoExcel As String = openFileDialog1.FileName

            ' Inicializa una aplicación Excel
            Dim excelApp As New Application()
            Dim excelLibro As Workbook = excelApp.Workbooks.Open(archivoExcel)
            Dim excelHoja As Worksheet = CType(excelLibro.Sheets(1), Worksheet)

            ' Obtiene el rango de celdas con datos
            Dim rango As Range = excelHoja.UsedRange

            ' Obtiene el número de filas y columnas en el rango
            Dim numRows As Integer = rango.Rows.Count
            Dim numCols As Integer = rango.Columns.Count

            ' Limpia el DataGridView antes de cargar nuevos datos
            DataGridView1.Rows.Clear()
            'DataGridView1.Columns.Clear()

            ' Agrega columnas al DataGridView
            'For j As Integer = 1 To numCols

            '    DataGridView1.Columns.Add(rango.Cells(2, j).Value, rango.Cells(2, j).Value)


            'Next

            ' Agrega filas al DataGridView
            For i As Integer = 2 To numRows
                Dim fila As New List(Of Object)()
                For j As Integer = 1 To numCols
                    fila.Add(rango.Cells(i, j).Value)
                Next
                DataGridView1.Rows.Add(fila.ToArray())
            Next

            ' Cierra y libera los recursos de Excel
            excelLibro.Close()
            excelApp.Quit()
            Marshal.ReleaseComObject(excelLibro)
            Marshal.ReleaseComObject(excelApp)




        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim p As Integer
        p = DataGridView1.Rows.Count
        abrir()
        For i = 0 To p - 1
            Dim cmd20 As New SqlCommand("insert into Listar_Fec_Doc(CueLiD,RucLiD,RSoLiD,CodLiD,NuDLiD,FecLiD,PrvLid) values (@CueLiD,@RucLiD,@RSoLiD,@CodLiD,@NuDLiD,@FecLiD,@PrvLid)", conx)
            cmd20.Parameters.AddWithValue("@CueLiD", Trim(DataGridView1.Rows(i).Cells(0).Value))
            cmd20.Parameters.AddWithValue("@RucLiD", Trim(DataGridView1.Rows(i).Cells(1).Value))
            cmd20.Parameters.AddWithValue("@RSoLiD", Trim(DataGridView1.Rows(i).Cells(2).Value))
            cmd20.Parameters.AddWithValue("@CodLiD", Trim(DataGridView1.Rows(i).Cells(3).Value))
            cmd20.Parameters.AddWithValue("@NuDLiD", Trim(DataGridView1.Rows(i).Cells(4).Value))
            cmd20.Parameters.AddWithValue("@FecLiD", Convert.ToDateTime(DataGridView1.Rows(i).Cells(5).Value).Date)
            cmd20.Parameters.AddWithValue("@PrvLid", Convert.ToInt32(DataGridView1.Rows(i).Cells(6).Value))
            cmd20.ExecuteNonQuery()
        Next
        MsgBox("SE REGISTRO CON EXITO")
        DataGridView1.Rows.Clear()
        'DataGridView1.Columns.Clear()
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        abrir()
        '        Dim sql1021 As String = "select CueLiD, RucLiD,RSoLiD,CodLiD,NuDLiD,FecLiD,PrvLid, isnull( isnull(c.fcom_3,p.fref_3a),'') AS FECHA,
        'ISNULL( isnull(c.nomb_3, p.nomb_3a),'') AS RSOCIAL from Listar_Fec_Doc  l 
        'left join custom_vianny.dbo.fag0300 c on l.RucLiD = c.fich_3 and l.NuDLiD= (c.sfactu_3 +'-'+ c.nfactu_3) and c.ccia_3 ='01'
        'left join custom_vianny.dbo.cac3p00 p on p.ndoc_3a = l.NuDLiD AND p.ccom_3a ='333' and p.ccia_3a ='01'"
        Dim sql1021 As String = "select CueLiD, RucLiD,RSoLiD,CodLiD,NuDLiD,FecLiD,PrvLid, isnull(fcom_3,'') AS FECHA,
        ISNULL(gloa_3,'') AS RSOCIAL from Listar_Fec_Doc  l 
        left join custom_vianny.dbo.fag0300 c on  l.NuDLiD= (c.sfactu_3 +'-'+ c.nfactu_3) and c.ccia_3 ='" + Label1.Text + "'"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr21 = cmd1021.ExecuteReader()
        Dim i51 As Integer
        i51 = 0

        While Rsr21.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i51).Cells(0).Value = Rsr21(0)
            DataGridView1.Rows(i51).Cells(1).Value = Rsr21(1)
            DataGridView1.Rows(i51).Cells(2).Value = Rsr21(2)
            DataGridView1.Rows(i51).Cells(3).Value = Rsr21(3)
            DataGridView1.Rows(i51).Cells(4).Value = Rsr21(4)
            DataGridView1.Rows(i51).Cells(5).Value = Rsr21(5)
            DataGridView1.Rows(i51).Cells(6).Value = Rsr21(6)
            DataGridView1.Rows(i51).Cells(7).Value = Rsr21(7)
            DataGridView1.Rows(i51).Cells(8).Value = Rsr21(8)
            i51 = i51 + 1
        End While
        Rsr21.Close()
        Dim exportar As New ExportarData
        exportar.llenarExcel(DataGridView1)
        DataGridView1.Rows.Clear()
        Dim agregar1 As String = "delete from Listar_Fec_Doc"
        ELIMINAR(agregar1)
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
End Class