Imports System.IO 'esta libreria nos va a servir para poder activar el commandialog
Imports Microsoft.Office.Interop
Imports System.Data
Imports System.Data.OleDb
Imports System
Imports Microsoft.VisualBasic
Module Importar_registros
    Sub importarExcel(ByVal tabla As DataGridView)
        Dim myFileDialog As New OpenFileDialog()
        Dim xSheet As String = ""

        With myFileDialog
            .Filter = "Excel Files |*.xlsm"
            .Title = "Open File"
            .ShowDialog()
        End With
        If myFileDialog.FileName.ToString <> "" Then
            Dim ExcelFile As String = myFileDialog.FileName.ToString

            Dim ds As New DataSet
            Dim da As OleDbDataAdapter
            Dim dt As DataTable
            Dim conn As OleDbConnection

            xSheet = "FORMATO"
            conn = New OleDbConnection(
                              "Provider=Microsoft.ACE.OLEDB.12.0;" &
                              "data source=" & ExcelFile & "; " &
                             "Extended Properties='Excel 12.0 Xml;HDR=Yes'")

            Try
                da = New OleDbDataAdapter("SELECT * FROM  [" & xSheet & "$]", conn)

                conn.Open()
                da.Fill(ds, "MyData")
                dt = limpiarFilas(ds.Tables("MyData"))
                dt = ds.Tables("MyData")
                tabla.DataSource = ds
                tabla.DataMember = "MyData"
                MsgBox("Se ha cargado la importacion correctamente", MsgBoxStyle.Information, "Importado con exito")

            Catch ex As Exception
                MsgBox(ex.Message)
                MsgBox("Inserte un nombre valido de la Hoja que desea importar", MsgBoxStyle.Information, "Informacion")
            Finally
                conn.Close()
            End Try
        End If
    End Sub
    Sub importarExcel2(ByVal tabla As DataGridView)
        Dim myFileDialog As New OpenFileDialog()
        Dim xSheet As String = ""

        With myFileDialog
            .Filter = "Excel Files |*.xlsx"
            .Title = "Open File"
            .ShowDialog()
        End With
        If myFileDialog.FileName.ToString <> "" Then
            Dim ExcelFile As String = myFileDialog.FileName.ToString

            Dim ds As New DataSet
            Dim da As OleDbDataAdapter
            Dim dt As DataTable
            Dim conn As OleDbConnection

            xSheet = "FORMATO"
            conn = New OleDbConnection(
                              "Provider=Microsoft.ACE.OLEDB.12.0;" &
                              "data source=" & ExcelFile & "; " &
                             "Extended Properties='Excel 12.0 Xml;HDR=Yes'")

            Try
                da = New OleDbDataAdapter("SELECT * FROM  [" & xSheet & "$]", conn)

                conn.Open()
                da.Fill(ds, "MyData")
                dt = limpiarFilas(ds.Tables("MyData"))
                dt = ds.Tables("MyData")
                tabla.DataSource = ds
                tabla.DataMember = "MyData"
                MsgBox("Se ha cargado la importacion correctamente", MsgBoxStyle.Information, "Importado con exito")

            Catch ex As Exception

                MsgBox("Inserte un nombre valido de la Hoja que desea importar", MsgBoxStyle.Information, "Informacion")
            Finally
                conn.Close()
            End Try
        End If
    End Sub

    Public Function limpiarFilas(ByVal tb As DataTable) As DataTable

        'obtenemos la cantidad de columnas del DataTable
        Dim columnas As Integer = tb.Columns.Count

        'Recorremos las filas
        For Each fila As DataRow In tb.Rows
            Dim vacios As Integer = 0
            For i As Integer = 0 To columnas - 1
                If String.IsNullOrEmpty(Convert.ToString(fila(i))) Then
                    vacios += 1
                End If
            Next

            If vacios = columnas Then
                fila.Delete()
            End If
        Next

        Return tb

    End Function
End Module
