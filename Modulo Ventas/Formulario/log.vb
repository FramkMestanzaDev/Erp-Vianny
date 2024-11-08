Imports System.Data.SqlClient
Public Class log
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub log_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "select *, case when accion  ='1' then 'CREAR' when accion  ='2' then 'MODIFICAR' when accion  ='3' then 'ELIMINAR' when accion  ='4' then 'APROBACION_COMERCIAL' when accion  ='5' then 'DESAPROBO_COMERCIAL' when accion  ='6' then 'APROBACION_COBRANZAS' when accion  ='7' then 'DESAPROBO_COBRANZAS' when accion  ='8' then 'APROBACION_ALMACEN' END from pllog where detalle ='" + Label1.Text + "' and modulo ='" + Label2.Text + "' and ccia ='" + Label3.Text + "' ORDER by fecha"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()

        Dim i51 As Integer
        While Rsr1991.Read()

            DataGridView1.Rows.Add()
            DataGridView1.Rows(i51).Cells(0).Value = Rsr1991(2)
            DataGridView1.Rows(i51).Cells(1).Value = Rsr1991(0)
            DataGridView1.Rows(i51).Cells(2).Value = Rsr1991(1)
            DataGridView1.Rows(i51).Cells(3).Value = Rsr1991(3)
            DataGridView1.Rows(i51).Cells(4).Value = Rsr1991(4)
            DataGridView1.Rows(i51).Cells(5).Value = Rsr1991(7)
            DataGridView1.Rows(i51).Cells(6).Value = Rsr1991(5)
            i51 = i51 + 1
        End While

        Rsr1991.Close()
    End Sub
End Class