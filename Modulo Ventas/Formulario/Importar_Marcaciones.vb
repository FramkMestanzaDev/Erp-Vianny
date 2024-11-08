Imports System.Data
Imports System.Data.OleDb
Public Class Importar_Marcaciones

    Dim cadena As New OleDbConnection
    Private Sub Importar_Marcaciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            cadena.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data  Data Source=C:\Program Files (x86)\ZKTeco\att2000.mdb"
            cadena.Open()

            Mostrar()
            MsgBox("Conectado con la Base de Datos", vbInformation, "Aviso")
        Catch ex As Exception
            MsgBox("No se conecto con la Base de Datos", vbCritical, "Aviso")
        End Try
    End Sub
    Private Sub Mostrar()

        Dim oda As New OleDbDataAdapter
        Dim ods As New DataSet
        Dim consulta As String

        consulta = "Select *From USERINFO"
        oda = New OleDbDataAdapter(consulta, cadena)
        ods.Tables.Add("USERINFO")
        oda.Fill(ods.Tables("Paciente"))
        DataGridView1.DataSource = ods.Tables("Paciente")
    End Sub
End Class