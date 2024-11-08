Imports System.Data.SqlClient
Public Class rbros_libre
    Public conx As SqlConnection
    Public conn As SqlDataAdapter

    Public comando As SqlCommand
    Dim da As New DataTable
    Dim bs As New BindingSource()
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub rbros_libre_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT oper_14f AS RUBRO,nomb_14f as DESCRIPCION FROM custom_vianny.DBO.fag1400 WHERE  ccia_14f='" + Label2.Text + "' AND VEN=1", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).Width = 350
            DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Registro_Ventas.TextBox16.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Registro_Ventas.TextBox17.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Registro_Ventas.TextBox13.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Registro_Ventas.TextBox24.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Registro_Ventas.Label26.Text = Label2.Text
        Registro_Ventas.Label17.Text = Label3.Text
        Registro_Ventas.ShowDialog()
    End Sub
End Class