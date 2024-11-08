Imports System.Data.SqlClient
Public Class ReporteMatPo
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
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
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            ListaPO.Label3.Text = "2"
            ListaPO.ShowDialog()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ReporteMatrizPo.TextBox1.Text = Label2.Text
        ReporteMatrizPo.TextBox2.Text = TextBox1.Text
        ReporteMatrizPo.ShowDialog()
    End Sub

    Private Sub ReporteMatPo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()
        If CheckBox1.Checked = True Then
            Dim dg2 As New DataTable
            dg2.Clear()
            Dim cmd2 As New SqlDataAdapter("Sp_ReporteMatrizConsumoPOExportar '" + Label2.Text + "','" + TextBox1.Text.ToString().Trim() + "','01'", conx)
            cmd2.Fill(dg2)
            If dg2.Rows.Count <> 0 Then
                DataGridView1.DataSource = dg2
            End If
            Dim export As New Exportar
            export.llenarExcel(DataGridView1)
        Else
            Dim dg2 As New DataTable
            dg2.Clear()
            Dim cmd2 As New SqlDataAdapter("Sp_ReporteMatrizConsumoConsolidadoPOExportar '" + Label2.Text + "','" + TextBox1.Text.ToString().Trim() + "','01'", conx)
            cmd2.Fill(dg2)
            If dg2.Rows.Count <> 0 Then
                DataGridView1.DataSource = dg2
            End If
            Dim export As New Exportar
            export.llenarExcel(DataGridView1)
        End If

    End Sub
End Class