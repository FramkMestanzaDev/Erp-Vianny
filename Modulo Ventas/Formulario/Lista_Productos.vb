Imports System.Data.SqlClient
Public Class Lista_Productos
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
    Dim DT14 As New DataTable
    Private Sub Lista_Productos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        DataGridView1.DataSource = ""

        DT14.Clear()
        Dim fec As String
        fec = Label5.Text & "1231"

        Dim cmd6 As New SqlDataAdapter("exec CaeSoft_ReporteStockFisico_ayuda_osc '" + Label2.Text + "' , '" + Label3.Text + "', NULL, '" + Label5.Text + "','" + fec + "', '" + Trim(Label4.Text) + "'", conx)
        cmd6.Fill(DT14)



        DataGridView1.DataSource = DT14
        DataGridView1.Columns(0).Width = 45
        DataGridView1.Columns(1).Width = 140
        DataGridView1.Columns(2).Visible = False
        DataGridView1.Columns(3).Width = 370
        DataGridView1.Columns(4).Width = 45
        DataGridView1.Columns(5).Width = 220

    End Sub
    Sub BUSCAR1()
        Dim ds As New DataSet
        ds.Tables.Add(DT14.Copy)
        Dim dv As New DataView(ds.Tables(0))
        dv.RowFilter = "PRODUCTO" & " like '%" & TextBox1.Text & "%'"
        If dv.Count <> 0 Then
            DataGridView1.Columns(0).Width = 45
            DataGridView1.Columns(1).Width = 140
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(3).Width = 370
            DataGridView1.Columns(4).Width = 45
            DataGridView1.Columns(5).Width = 220

        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        BUSCAR1()
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Ocs.DataGridView1.Rows(Label7.Text).Cells(4).Value = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Ocs.DataGridView1.Rows(Label7.Text).Cells(5).Value = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        Me.Close()
    End Sub
End Class