Imports System.Data.SqlClient
Public Class RptOpPrecoso
    Dim da1 As New DataTable
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            conx.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        mostar_todos()
        Dim EXP As New Exportar
        EXP.llenarExcel(DataGridView1)
    End Sub
    Private Sub mostar_todos()
        abrir()
        da1.Clear()
        DataGridView1.DataSource = ""

        Dim cmd As New SqlDataAdapter("EXEC Sp_Reporte_Cortes_Asignados '01','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "','" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "',NULL,NULL,NULL,NULL,1,1", conx)
        cmd.Fill(da1)
        DataGridView1.DataSource = da1
    End Sub
End Class