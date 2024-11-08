Imports System.Data.SqlClient
Public Class Otras_Facturas
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
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
    Dim bs As New BindingSource()

    Dim da As New DataTable

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = ""
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select empresa ,ano,ncom AS COMPROBANTE,ruc AS RUC,documento AS DOCUMENTO,cuenta AS CUENTA,provision AS PROVISION,sigla_voucher AS VOUCHER,observacion,estado,id_feliminadas from facturas_eliminadas where estado=" + Label4.Text + "", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        Dim y As Integer

        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(11).Visible = False
            y = DataGridView1.Rows.Count
            For i = 0 To y - 1
                DataGridView1.Rows(i).Cells(0).Value = Trim(DataGridView1.Rows(i).Cells(9).Value)
            Next
        End If

    End Sub




    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex = -1 Then
            MsgBox("Esta Funcion esta Bloqueada")
            For Each col As DataGridViewColumn In DataGridView1.Columns
                col.SortMode = DataGridViewColumnSortMode.NotSortable
            Next
        Else
            If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OBSERVACION" Then
                OBSERVACION.Label2.Text = 1
                OBSERVACION.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                OBSERVACION.Label3.Text = e.RowIndex
                OBSERVACION.Label4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value)
                OBSERVACION.ShowDialog()
            End If
        End If
    End Sub
End Class