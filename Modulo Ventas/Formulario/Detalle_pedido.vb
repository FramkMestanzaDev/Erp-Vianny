Imports System.Data.SqlClient
Public Class Detalle_pedido
    Dim DT, DT2 As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim ty, ty2, ty3 As New DataTable
    Public conn As SqlDataAdapter

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Dim Rsr22 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Detalle_pedido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim i10 As Integer
        i10 = 0
        Dim sql1022 As String = "select codigo_tela,descripcion,partida,cantidad from Comercial_Textil.DBO.detalle_pedido where numero_pedido ='" + Trim(Label1.Text) + "' and ESTADO ='1'"
        Dim cmd1022 As New SqlCommand(sql1022, conx)
        Rsr22 = cmd1022.ExecuteReader()
        While Rsr22.Read
            DataGridView1.Rows.Add()

            DataGridView1.Rows(i10).Cells(0).Value = Trim(Rsr22(0))
            DataGridView1.Rows(i10).Cells(1).Value = Trim(Rsr22(1))
            DataGridView1.Rows(i10).Cells(2).Value = Trim(Rsr22(2))
            DataGridView1.Rows(i10).Cells(3).Value = Trim(Rsr22(3))
            i10 = i10 + 1
        End While
        Rsr22.Close()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Hoja_Reclamos.DataGridView1.Rows(Label2.Text).Cells(1).Value = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Hoja_Reclamos.DataGridView1.Rows(Label2.Text).Cells(2).Value = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Hoja_Reclamos.DataGridView1.Rows(Label2.Text).Cells(3).Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        Hoja_Reclamos.DataGridView1.Rows(Label2.Text).Cells(4).Value = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        Me.Close()
    End Sub
End Class