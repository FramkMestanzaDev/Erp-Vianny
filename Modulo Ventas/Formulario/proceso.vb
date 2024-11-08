Imports System.Data.SqlClient
Public Class proceso
    Public conx As SqlConnection
    Dim dg1 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dg1.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 300
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub proceso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT fase_4 AS CODIGO,nomb_4 AS DESCRIPCION FROM custom_vianny.dbo.Qag0400 A  Where CCIA = '01'", conx)
        cmd.Fill(dg1)
        If dg1.Rows.Count <> 0 Then
            DataGridView1.DataSource = dg1
            DataGridView1.Columns(1).Width = 300
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If Label2.Text = 1 Then
            Guia_hilo.TextBox24.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            Guia_hilo.TextBox23.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
            Me.Close()
        Else
            Guia_remision.TextBox24.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            Guia_remision.TextBox23.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
            Me.Close()
        End If

    End Sub
End Class