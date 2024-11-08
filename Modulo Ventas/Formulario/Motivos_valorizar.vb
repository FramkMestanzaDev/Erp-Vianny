Imports System.Data.SqlClient
Public Class Motivos_valorizar
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
    Private Sub Motivos_valorizar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        da.Clear()
        DataGridView1.DataSource = Nothing
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT dest_21e AS CODIGO,nomb_21e AS DESCRIPCION FROM  custom_vianny.dbo.Cag210e A  Where A.empr_21e + A.almac_21e ='" + Label1.Text & Label2.Text + "'", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).Width = 300
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
    End Sub

    Private Sub Motivos_valorizar_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim p As Integer
        Dim motivo, motivof As String
        motivo = ""
        motivof = ""
        p = DataGridView1.Rows.Count
        For i = 0 To p - 1

            motivo = DataGridView1.Rows(i).Cells(0).Value
            If i = 0 Then
                motivof = motivo
            Else
                motivof = motivof + "," + motivo
            End If


        Next
        If Label3.Text = 1 Then
            Valorizr_Salidas.TextBox1.Text = Trim(motivof)
        Else
            Valorizr_Salidas.TextBox5.Text = Trim(motivof)
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class