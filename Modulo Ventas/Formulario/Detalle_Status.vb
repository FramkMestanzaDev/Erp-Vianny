Imports System.Data.SqlClient
Public Class Detalle_Status
    Public conx As SqlConnection
    Public comando As SqlCommand
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

    Dim da As New DataTable
    Sub mostrar_informacion()
        abrir()
        da.Clear()
        DataGridView1.DataSource = Nothing
        Dim cmd As New SqlDataAdapter("SELECT NomRut as PROCESO,fechRut as FEcha, AsiRut as ASIGNACION, ObsRut as OBSERVACION FROM  Ruta_Udp WHERE OdRut ='" + Trim(TextBox1.Text) + "'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 180
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(2).Width = 250
        DataGridView1.Columns(3).Width = 250
        Dim P As Integer
        P = 0
        P = DataGridView1.Rows.Count
        If P > 0 Then
            For I = 0 To 7
                If DateTime.Compare(DateTimePicker1.Value, DataGridView1.Rows(I).Cells(1).Value) > 0 Then
                    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.Red
                    DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White
                Else
                    DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.White
                    DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.Black
                End If
            Next
        End If


    End Sub
    Private Sub Detalle_Status_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mostrar_informacion()
    End Sub
End Class